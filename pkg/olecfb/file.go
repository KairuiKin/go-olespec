package olecfb

import (
	"bytes"
	"context"
	"errors"
	"io"
	"os"
	"strings"
	"sync"

	"github.com/KairuiKin/go-olespec/pkg/olecfb/storage"
)

type File struct {
	rb       storage.ReadBackend
	wb       storage.WriteBackend
	opt      OpenOptions
	closed   bool
	root     Node
	hdr      *cfbHeader
	fat      []uint32
	miniFAT  []uint32
	miniData []byte
	nodes    map[NodeID]Node
	order    []NodeID
	children map[NodeID][]NodeID
	entries  map[NodeID]dirEntry
	report   Report
	mu       sync.RWMutex
}

type Tx struct {
	file   *File
	opt    TxOptions
	closed bool
}

func Open(r storage.ReadBackend, opt OpenOptions) (*File, error) {
	if r == nil {
		return nil, newError(ErrInvalidArgument, "read backend is nil", "open", "", -1, nil)
	}
	hdr, err := parseHeader(r.ReadAt, r.Size())
	if err != nil {
		return nil, err
	}
	report := Report{Mode: opt.Mode}

	fat, err := loadFAT(r.ReadAt, r.Size(), hdr)
	if err != nil {
		if opt.Mode == Lenient {
			report.Partial = true
			report.Warnings = append(report.Warnings, warningFromError(err, SeverityWarning))
			fat = nil
		} else {
			return nil, err
		}
	}
	root, err := parseRootNode(r.ReadAt, r.Size(), hdr)
	if err != nil {
		if opt.Mode == Lenient {
			report.Partial = true
			report.Warnings = append(report.Warnings, warningFromError(err, SeverityWarning))
			root = Node{
				ID:   0,
				Type: NodeRoot,
				Path: "/",
				Name: "Root Entry",
			}
		} else {
			return nil, err
		}
	}

	entries, err := parseDirectoryEntries(r.ReadAt, r.Size(), hdr, fat, opt.MaxChainLength)
	if err != nil {
		if opt.Mode == Lenient {
			report.Partial = true
			report.Warnings = append(report.Warnings, warningFromError(err, SeverityWarning))
		} else {
			return nil, err
		}
	}

	nodes := map[NodeID]Node{0: root}
	order := []NodeID{0}
	entryMap := map[NodeID]dirEntry{}
	for _, e := range entries {
		entryMap[NodeID(e.ID)] = e
	}
	if len(entries) > 0 {
		treeRoot, parsedNodes, parsedOrder, treeErr := buildNodesFromEntries(entries, opt.MaxRecursion, opt.MaxObjectCount)
		if treeErr != nil {
			if opt.Mode == Lenient {
				report.Partial = true
				report.Warnings = append(report.Warnings, warningFromError(treeErr, SeverityWarning))
			} else {
				return nil, treeErr
			}
		} else {
			root = treeRoot
			nodes = parsedNodes
			order = parsedOrder
		}
	}

	miniFAT, miniData, miniErr := loadMiniData(r.ReadAt, r.Size(), hdr, fat, entryMap, opt.MaxChainLength)
	if miniErr != nil {
		if opt.Mode == Lenient {
			report.Partial = true
			report.Warnings = append(report.Warnings, warningFromError(miniErr, SeverityWarning))
			miniFAT = nil
			miniData = nil
		} else {
			return nil, miniErr
		}
	}

	f := &File{
		rb:       r,
		opt:      opt,
		root:     root,
		hdr:      hdr,
		fat:      fat,
		miniFAT:  miniFAT,
		miniData: miniData,
		nodes:    nodes,
		order:    order,
		children: buildChildrenIndex(order, nodes),
		entries:  entryMap,
		report:   report,
	}
	return f, nil
}

func OpenReadWrite(rw storage.WriteBackend, opt OpenOptions) (*File, error) {
	if rw == nil {
		return nil, newError(ErrInvalidArgument, "write backend is nil", "open_rw", "", -1, nil)
	}
	f, err := Open(rw, opt)
	if err != nil {
		return nil, err
	}
	f.wb = rw
	return f, nil
}

func OpenFile(path string, opt OpenOptions) (*File, error) {
	fp, err := os.Open(path)
	if err != nil {
		return nil, newError(ErrNotFound, "open file failed", "open_file", path, -1, err)
	}
	return Open(newFileBackend(path, fp, true), opt)
}

func OpenBytes(buf []byte, opt OpenOptions) (*File, error) {
	return Open(newMemBackend("mem://readonly", append([]byte(nil), buf...), true), opt)
}

func OpenBytesRW(buf []byte, opt OpenOptions) (*File, error) {
	return OpenReadWrite(newMemBackend("mem://rw", append([]byte(nil), buf...), false), opt)
}

func CreateInMemory(opt CreateOptions) (*File, error) {
	backend := newMemBackend("mem://new", nil, false)
	return newEmptyFile(backend, backend, OpenOptions{Mode: Strict})
}

func CreateFile(path string, opt CreateOptions) (*File, error) {
	fp, err := os.OpenFile(path, os.O_CREATE|os.O_RDWR|os.O_TRUNC, 0o644)
	if err != nil {
		return nil, newError(ErrInvalidArgument, "create file failed", "create_file", path, -1, err)
	}
	backend := newFileBackend(path, fp, false)
	return newEmptyFile(backend, backend, OpenOptions{Mode: Strict})
}

func (f *File) Begin(opt TxOptions) (*Tx, error) {
	if f == nil {
		return nil, newError(ErrInvalidArgument, "file is nil", "tx.begin", "", -1, nil)
	}
	if f.wb == nil {
		return nil, newError(ErrReadOnly, "file is read-only", "tx.begin", "", -1, nil)
	}
	if f.closed {
		return nil, newError(ErrInvalidArgument, "file is closed", "tx.begin", "", -1, nil)
	}
	return &Tx{file: f, opt: opt}, nil
}

func (f *File) Walk(fn func(Node) error) error {
	if f == nil {
		return newError(ErrInvalidArgument, "file is nil", "walk", "", -1, nil)
	}
	if fn == nil {
		return newError(ErrInvalidArgument, "walk callback is nil", "walk", "", -1, nil)
	}
	for _, id := range f.order {
		node, ok := f.nodes[id]
		if !ok {
			continue
		}
		if err := fn(node); err != nil {
			return err
		}
	}
	return nil
}

func (f *File) WalkEx(opt WalkOptions, fn func(WalkEvent) error) (WalkResult, error) {
	if fn == nil {
		return WalkResult{}, newError(ErrInvalidArgument, "walk callback is nil", "walk_ex", "", -1, nil)
	}
	sequence := f.order
	if opt.Order == WalkBFS {
		sequence = f.bfsOrder()
	}
	visited := 0
	maxDepth := 0
	for idx, id := range sequence {
		node, ok := f.nodes[id]
		if !ok {
			continue
		}
		depth := f.nodeDepth(id)
		if !opt.IncludeRoot && id == 0 {
			continue
		}
		if opt.MaxDepth > 0 && depth > opt.MaxDepth {
			continue
		}
		if depth > maxDepth {
			maxDepth = depth
		}
		event := WalkEvent{Node: node, Depth: depth, Index: idx}
		if err := fn(event); err != nil {
			return WalkResult{}, err
		}
		visited++
	}
	return WalkResult{Visited: visited, MaxDepth: maxDepth}, nil
}

func (f *File) OpenStream(path string) (StreamReader, error) {
	if f == nil {
		return nil, newError(ErrInvalidArgument, "file is nil", "open_stream", path, -1, nil)
	}
	if path == "" {
		return nil, newError(ErrInvalidArgument, "path is empty", "open_stream", path, -1, nil)
	}
	canonical, err := CanonicalPath(path)
	if err != nil {
		return nil, err
	}

	var target Node
	found := false
	for _, id := range f.order {
		n, ok := f.nodes[id]
		if !ok {
			continue
		}
		if strings.EqualFold(n.Path, string(canonical)) {
			target = n
			found = true
			break
		}
	}
	if !found {
		return nil, newError(ErrNotFound, "stream not found", "open_stream", string(canonical), -1, nil)
	}
	if target.Type != NodeStream {
		return nil, newError(ErrInvalidArgument, "path is not a stream", "open_stream", string(canonical), -1, nil)
	}
	if f.hdr == nil {
		return nil, newError(ErrBadHeader, "header is missing", "open_stream", string(canonical), -1, nil)
	}
	if target.Size < 0 {
		return nil, newError(ErrDirCorrupt, "negative stream size", "open_stream", string(canonical), -1, nil)
	}

	entry, ok := f.entries[target.ID]
	if !ok {
		return nil, newError(ErrDirCorrupt, "missing stream directory entry", "open_stream", string(canonical), -1, nil)
	}

	if uint64(target.Size) < uint64(f.hdr.MiniStreamCutoff) {
		if target.Size == 0 {
			return newBytesStreamReader(target, nil), nil
		}
		data, err := readMiniStreamData(
			f.miniData,
			f.miniFAT,
			f.hdr,
			entry.StartSector,
			target.Size,
			f.opt.MaxChainLength,
		)
		if err != nil {
			return nil, err
		}
		return newBytesStreamReader(target, data), nil
	}
	if len(f.fat) == 0 {
		return nil, newError(ErrBadFATChain, "fat table is empty", "open_stream", string(canonical), -1, nil)
	}
	if entry.StartSector == cfbEndOfChain {
		if target.Size == 0 {
			return newBytesStreamReader(target, nil), nil
		}
		return nil, newError(ErrBadFATChain, "stream has data size but no sector chain", "open_stream", string(canonical), -1, nil)
	}
	data, err := readNormalStreamData(f.rb.ReadAt, f.rb.Size(), f.fat, f.hdr, entry.StartSector, target.Size, f.opt.MaxChainLength)
	if err != nil {
		return nil, err
	}
	return newBytesStreamReader(target, data), nil
}

func (f *File) SnapshotBytes() ([]byte, error) {
	if f == nil || f.rb == nil {
		return nil, newError(ErrInvalidArgument, "file is not initialized", "snapshot", "", -1, nil)
	}
	ib, ok := f.rb.(storage.Introspect)
	if !ok || ib.Info().Kind != storage.KindMem {
		return nil, newError(ErrUnsupported, "snapshot is available for mem backend only", "snapshot", "", -1, nil)
	}
	mb, ok := f.rb.(*memBackend)
	if !ok {
		return nil, newError(ErrUnsupported, "unsupported memory backend", "snapshot", "", -1, nil)
	}
	return mb.snapshot(), nil
}

func (f *File) Close() error {
	if f == nil || f.closed {
		return nil
	}
	f.closed = true
	if f.rb != nil {
		return f.rb.Close()
	}
	return nil
}

func (f *File) Report() Report {
	if f == nil {
		return Report{}
	}
	f.mu.RLock()
	defer f.mu.RUnlock()
	out := f.report
	if len(f.report.Warnings) > 0 {
		out.Warnings = append([]Warning(nil), f.report.Warnings...)
	}
	if len(f.report.Repairs) > 0 {
		out.Repairs = append([]RepairRecord(nil), f.report.Repairs...)
	}
	return out
}

func (f *File) GetNode(id NodeID) (Node, error) {
	if f == nil {
		return Node{}, newError(ErrInvalidArgument, "file is nil", "node.get", "", -1, nil)
	}
	n, ok := f.nodes[id]
	if !ok {
		return Node{}, newError(ErrNotFound, "node not found", "node.get", "", -1, nil)
	}
	return n, nil
}

func (f *File) GetNodeByPath(path string) (Node, error) {
	if f == nil {
		return Node{}, newError(ErrInvalidArgument, "file is nil", "node.get_by_path", path, -1, nil)
	}
	canonical, err := CanonicalPath(path)
	if err != nil {
		return Node{}, err
	}
	for _, id := range f.order {
		n, ok := f.nodes[id]
		if !ok {
			continue
		}
		if strings.EqualFold(n.Path, string(canonical)) {
			return n, nil
		}
	}
	return Node{}, newError(ErrNotFound, "node path not found", "node.get_by_path", string(canonical), -1, nil)
}

func (f *File) ListNodes() []Node {
	if f == nil {
		return nil
	}
	out := make([]Node, 0, len(f.order))
	for _, id := range f.order {
		n, ok := f.nodes[id]
		if !ok {
			continue
		}
		out = append(out, n)
	}
	return out
}

func (f *File) nodeDepth(id NodeID) int {
	node, ok := f.nodes[id]
	if !ok {
		return 0
	}
	depth := 0
	guard := 0
	for node.ID != 0 {
		parent, ok := f.nodes[node.ParentID]
		if !ok {
			break
		}
		node = parent
		depth++
		guard++
		if guard > len(f.nodes) {
			break
		}
	}
	return depth
}

func (f *File) bfsOrder() []NodeID {
	if len(f.nodes) == 0 {
		return nil
	}
	out := make([]NodeID, 0, len(f.nodes))
	queue := []NodeID{0}
	seen := map[NodeID]struct{}{0: {}}
	for len(queue) > 0 {
		id := queue[0]
		queue = queue[1:]
		out = append(out, id)
		for _, child := range f.children[id] {
			if _, ok := seen[child]; ok {
				continue
			}
			seen[child] = struct{}{}
			queue = append(queue, child)
		}
	}
	return out
}

func (tx *Tx) PutStream(path string, r io.Reader, size int64) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.put_stream", path, -1, nil)
	}
	if r == nil {
		return newError(ErrInvalidArgument, "reader is nil", "tx.put_stream", path, -1, nil)
	}
	return newError(ErrUnsupported, "write path is not implemented yet", "tx.put_stream", path, -1, nil)
}

func (tx *Tx) Delete(path string) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.delete", path, -1, nil)
	}
	return newError(ErrUnsupported, "delete is not implemented yet", "tx.delete", path, -1, nil)
}

func (tx *Tx) Rename(oldPath, newPath string) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.rename", oldPath, -1, nil)
	}
	return newError(ErrUnsupported, "rename is not implemented yet", "tx.rename", oldPath, -1, nil)
}

func (tx *Tx) CreateStorage(path string) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.create_storage", path, -1, nil)
	}
	return newError(ErrUnsupported, "create storage is not implemented yet", "tx.create_storage", path, -1, nil)
}

func (tx *Tx) Commit(ctx context.Context, opt CommitOptions) (*CommitResult, error) {
	if tx == nil || tx.closed {
		return nil, newError(ErrTxClosed, "transaction is closed", "tx.commit", "", -1, nil)
	}
	if ctx == nil {
		ctx = context.Background()
	}
	tx.closed = true
	if errors.Is(ctx.Err(), context.Canceled) {
		return nil, newError(ErrCommitFailed, "commit canceled", "tx.commit", "", -1, ctx.Err())
	}
	if errors.Is(ctx.Err(), context.DeadlineExceeded) {
		return nil, newError(ErrCommitFailed, "commit timed out", "tx.commit", "", -1, ctx.Err())
	}
	if tx.file.wb != nil && opt.Sync {
		if err := tx.file.wb.Sync(); err != nil {
			return nil, newError(ErrCommitFailed, "sync failed", "tx.commit", "", -1, err)
		}
	}
	size := tx.file.rb.Size()
	return &CommitResult{
		BytesWritten: 0,
		NewSize:      size,
		BackendKind:  backendKind(tx.file.rb),
	}, nil
}

func (tx *Tx) Revert() error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.revert", "", -1, nil)
	}
	tx.closed = true
	return nil
}

func backendKind(rb storage.ReadBackend) string {
	if rb == nil {
		return ""
	}
	introspect, ok := rb.(storage.Introspect)
	if !ok {
		return ""
	}
	return string(introspect.Info().Kind)
}

func warningFromError(err error, severity Severity) Warning {
	var oe *OLEError
	if errors.As(err, &oe) {
		return Warning{
			Code:     oe.Code,
			Message:  oe.Message,
			Path:     oe.Path,
			Offset:   oe.Offset,
			Op:       oe.Op,
			Severity: severity,
		}
	}
	return Warning{
		Code:     ErrUnsupported,
		Message:  err.Error(),
		Offset:   -1,
		Severity: severity,
	}
}

func newEmptyFile(rb storage.ReadBackend, wb storage.WriteBackend, opt OpenOptions) (*File, error) {
	root := Node{
		ID:   0,
		Type: NodeRoot,
		Path: "/",
		Name: "Root Entry",
	}
	return &File{
		rb:       rb,
		wb:       wb,
		opt:      opt,
		root:     root,
		nodes:    map[NodeID]Node{0: root},
		order:    []NodeID{0},
		children: map[NodeID][]NodeID{},
		entries:  map[NodeID]dirEntry{},
	}, nil
}

func buildChildrenIndex(order []NodeID, nodes map[NodeID]Node) map[NodeID][]NodeID {
	children := make(map[NodeID][]NodeID, len(nodes))
	for _, id := range order {
		if id == 0 {
			continue
		}
		n, ok := nodes[id]
		if !ok {
			continue
		}
		children[n.ParentID] = append(children[n.ParentID], id)
	}
	return children
}

type fileBackend struct {
	name     string
	file     *os.File
	readOnly bool
}

func newFileBackend(name string, file *os.File, readOnly bool) *fileBackend {
	return &fileBackend{name: name, file: file, readOnly: readOnly}
}

func (b *fileBackend) ReadAt(p []byte, off int64) (n int, err error) { return b.file.ReadAt(p, off) }
func (b *fileBackend) Size() int64 {
	st, err := b.file.Stat()
	if err != nil {
		return 0
	}
	return st.Size()
}
func (b *fileBackend) Close() error { return b.file.Close() }
func (b *fileBackend) WriteAt(p []byte, off int64) (n int, err error) {
	if b.readOnly {
		return 0, newError(ErrReadOnly, "file backend is read-only", "backend.write_at", b.name, off, nil)
	}
	return b.file.WriteAt(p, off)
}
func (b *fileBackend) Truncate(size int64) error {
	if b.readOnly {
		return newError(ErrReadOnly, "file backend is read-only", "backend.truncate", b.name, -1, nil)
	}
	return b.file.Truncate(size)
}
func (b *fileBackend) Sync() error {
	if b.readOnly {
		return nil
	}
	return b.file.Sync()
}
func (b *fileBackend) Info() storage.Info {
	return storage.Info{
		Kind:     storage.KindFile,
		ReadOnly: b.readOnly,
		Name:     b.name,
	}
}

type memBackend struct {
	name     string
	readOnly bool
	mu       sync.RWMutex
	buf      []byte
}

func newMemBackend(name string, buf []byte, readOnly bool) *memBackend {
	return &memBackend{name: name, buf: buf, readOnly: readOnly}
}

func (b *memBackend) ReadAt(p []byte, off int64) (int, error) {
	b.mu.RLock()
	defer b.mu.RUnlock()
	if off < 0 {
		return 0, io.EOF
	}
	if off >= int64(len(b.buf)) {
		return 0, io.EOF
	}
	n := copy(p, b.buf[off:])
	if n < len(p) {
		return n, io.EOF
	}
	return n, nil
}

func (b *memBackend) WriteAt(p []byte, off int64) (int, error) {
	b.mu.Lock()
	defer b.mu.Unlock()
	if b.readOnly {
		return 0, newError(ErrReadOnly, "mem backend is read-only", "backend.write_at", b.name, off, nil)
	}
	if off < 0 {
		return 0, newError(ErrInvalidArgument, "negative offset", "backend.write_at", b.name, off, nil)
	}
	end := off + int64(len(p))
	if end > int64(len(b.buf)) {
		newBuf := make([]byte, end)
		copy(newBuf, b.buf)
		b.buf = newBuf
	}
	return copy(b.buf[off:end], p), nil
}

func (b *memBackend) Truncate(size int64) error {
	b.mu.Lock()
	defer b.mu.Unlock()
	if b.readOnly {
		return newError(ErrReadOnly, "mem backend is read-only", "backend.truncate", b.name, -1, nil)
	}
	if size < 0 {
		return newError(ErrInvalidArgument, "negative size", "backend.truncate", b.name, -1, nil)
	}
	if size <= int64(len(b.buf)) {
		b.buf = b.buf[:size]
		return nil
	}
	padding := make([]byte, size-int64(len(b.buf)))
	b.buf = append(b.buf, padding...)
	return nil
}

func (b *memBackend) Sync() error  { return nil }
func (b *memBackend) Close() error { return nil }

func (b *memBackend) Size() int64 {
	b.mu.RLock()
	defer b.mu.RUnlock()
	return int64(len(b.buf))
}

func (b *memBackend) snapshot() []byte {
	b.mu.RLock()
	defer b.mu.RUnlock()
	return bytes.Clone(b.buf)
}

func (b *memBackend) Info() storage.Info {
	return storage.Info{
		Kind:     storage.KindMem,
		ReadOnly: b.readOnly,
		Name:     b.name,
	}
}
