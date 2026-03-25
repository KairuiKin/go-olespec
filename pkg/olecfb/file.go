package olecfb

import (
	"bytes"
	"context"
	"errors"
	"io"
	"os"
	"path"
	"sort"
	"strings"
	"sync"

	"github.com/KairuiKin/go-olespec/pkg/olecfb/storage"
)

type File struct {
	rb       storage.ReadBackend
	wb       storage.WriteBackend
	opt      OpenOptions
	closed   bool
	activeTx *Tx
	root     Node
	hdr      *cfbHeader
	fat      []uint32
	miniFAT  []uint32
	miniData []byte

	nodes      map[NodeID]Node
	order      []NodeID
	children   map[NodeID][]NodeID
	entries    map[NodeID]dirEntry
	streamData map[NodeID][]byte
	report     Report

	mu sync.RWMutex
}

type Tx struct {
	file       *File
	opt        TxOptions
	closed     bool
	topologyChanged bool
	nodes      map[NodeID]Node
	order      []NodeID
	children   map[NodeID][]NodeID
	entries    map[NodeID]dirEntry
	streamData map[NodeID][]byte
	touchedStreams map[NodeID]struct{}
	nextID     NodeID
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
		rb:         r,
		opt:        opt,
		root:       root,
		hdr:        hdr,
		fat:        fat,
		miniFAT:    miniFAT,
		miniData:   miniData,
		nodes:      nodes,
		order:      order,
		children:   buildChildrenIndex(order, nodes),
		entries:    entryMap,
		streamData: map[NodeID][]byte{},
		report:     report,
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
	f.mu.Lock()
	defer f.mu.Unlock()
	if f.activeTx != nil {
		if f.activeTx.closed {
			f.activeTx = nil
		} else {
			return nil, newError(ErrConflict, "another transaction is active", "tx.begin", "", -1, nil)
		}
	}
	tx := &Tx{
		file:       f,
		opt:        opt,
		nodes:      cloneNodes(f.nodes),
		order:      append([]NodeID(nil), f.order...),
		children:   cloneChildren(f.children),
		entries:    cloneEntries(f.entries),
		streamData: cloneStreamData(f.streamData),
		touchedStreams: map[NodeID]struct{}{},
		nextID:     maxNodeID(f.nodes) + 1,
	}
	f.activeTx = tx
	return tx, nil
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

func (f *File) OpenStream(streamPath string) (StreamReader, error) {
	if f == nil {
		return nil, newError(ErrInvalidArgument, "file is nil", "open_stream", streamPath, -1, nil)
	}
	if streamPath == "" {
		return nil, newError(ErrInvalidArgument, "path is empty", "open_stream", streamPath, -1, nil)
	}
	canonical, err := CanonicalPath(streamPath)
	if err != nil {
		return nil, err
	}

	target, err := f.GetNodeByPath(string(canonical))
	if err != nil {
		return nil, err
	}
	if target.Type != NodeStream {
		return nil, newError(ErrInvalidArgument, "path is not a stream", "open_stream", string(canonical), -1, nil)
	}
	if data, ok := f.streamData[target.ID]; ok {
		return newBytesStreamReader(target, data), nil
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
		data, err := readMiniStreamData(f.miniData, f.miniFAT, f.hdr, entry.StartSector, target.Size, f.opt.MaxChainLength)
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

func (f *File) GetNodeByPath(p string) (Node, error) {
	if f == nil {
		return Node{}, newError(ErrInvalidArgument, "file is nil", "node.get_by_path", p, -1, nil)
	}
	canonical, err := CanonicalPath(p)
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

func (tx *Tx) PutStream(streamPath string, r io.Reader, size int64) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.put_stream", streamPath, -1, nil)
	}
	if r == nil {
		return newError(ErrInvalidArgument, "reader is nil", "tx.put_stream", streamPath, -1, nil)
	}
	data, err := io.ReadAll(r)
	if err != nil {
		return newError(ErrInvalidArgument, "failed to read stream data", "tx.put_stream", streamPath, -1, err)
	}
	if size >= 0 && int64(len(data)) != size {
		return newError(ErrInvalidArgument, "stream size does not match declared size", "tx.put_stream", streamPath, -1, nil)
	}
	p, err := CanonicalPath(streamPath)
	if err != nil {
		return err
	}
	parentPath := string(ParentPath(p))
	parentID, parentNode, ok := tx.findNodeByPath(parentPath)
	if !ok {
		return newError(ErrNotFound, "parent path not found", "tx.put_stream", parentPath, -1, nil)
	}
	if !parentNode.IsStorage() {
		return newError(ErrInvalidArgument, "parent is not a storage", "tx.put_stream", parentPath, -1, nil)
	}

	if id, node, exists := tx.findNodeByPath(string(p)); exists {
		if node.Type != NodeStream {
			return newError(ErrConflict, "path exists and is not a stream", "tx.put_stream", string(p), -1, nil)
		}
		if node.Size != int64(len(data)) {
			tx.topologyChanged = true
		}
		node.Size = int64(len(data))
		tx.nodes[id] = node
		e := tx.entries[id]
		e.Size = int64(len(data))
		tx.entries[id] = e
		tx.streamData[id] = append([]byte(nil), data...)
		tx.touchedStreams[id] = struct{}{}
		return nil
	}

	if tx.file.opt.MaxObjectCount > 0 && len(tx.nodes)+1 > tx.file.opt.MaxObjectCount {
		return newError(ErrLimitExceeded, "object count exceeded limit", "tx.put_stream", string(p), -1, nil)
	}
	id := tx.allocateID()
	name := pathBase(string(p))
	n := Node{
		ID:       id,
		Type:     NodeStream,
		Path:     string(p),
		Name:     name,
		ParentID: parentID,
		Size:     int64(len(data)),
	}
	tx.nodes[id] = n
	tx.entries[id] = dirEntry{
		ID:           uint32(id),
		Name:         name,
		ObjectType:   2,
		LeftSibling:  cfbNoStream,
		RightSibling: cfbNoStream,
		Child:        cfbNoStream,
		StartSector:  cfbEndOfChain,
		Size:         int64(len(data)),
	}
	tx.streamData[id] = append([]byte(nil), data...)
	tx.topologyChanged = true
	tx.children[parentID] = append(tx.children[parentID], id)
	parentNode.ChildCount = len(tx.children[parentID])
	tx.nodes[parentID] = parentNode
	tx.rebuildOrder()
	return nil
}

func (tx *Tx) Delete(deletePath string) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.delete", deletePath, -1, nil)
	}
	p, err := CanonicalPath(deletePath)
	if err != nil {
		return err
	}
	if p == "/" {
		return newError(ErrInvalidArgument, "cannot delete root", "tx.delete", string(p), -1, nil)
	}
	id, node, ok := tx.findNodeByPath(string(p))
	if !ok {
		return newError(ErrNotFound, "path not found", "tx.delete", string(p), -1, nil)
	}
	toDelete := tx.collectSubtree(id)
	tx.topologyChanged = true
	parentID := node.ParentID
	for _, did := range toDelete {
		delete(tx.nodes, did)
		delete(tx.entries, did)
		delete(tx.streamData, did)
		delete(tx.children, did)
	}
	tx.children[parentID] = filterNodeID(tx.children[parentID], toDelete)
	if parent, ok := tx.nodes[parentID]; ok {
		parent.ChildCount = len(tx.children[parentID])
		tx.nodes[parentID] = parent
	}
	tx.rebuildOrder()
	return nil
}

func (tx *Tx) Rename(oldPath, newPath string) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.rename", oldPath, -1, nil)
	}
	oldP, err := CanonicalPath(oldPath)
	if err != nil {
		return err
	}
	newP, err := CanonicalPath(newPath)
	if err != nil {
		return err
	}
	if oldP == "/" {
		return newError(ErrInvalidArgument, "cannot rename root", "tx.rename", string(oldP), -1, nil)
	}
	id, node, ok := tx.findNodeByPath(string(oldP))
	if !ok {
		return newError(ErrNotFound, "old path not found", "tx.rename", string(oldP), -1, nil)
	}
	if _, _, exists := tx.findNodeByPath(string(newP)); exists {
		return newError(ErrConflict, "new path already exists", "tx.rename", string(newP), -1, nil)
	}
	parentPath := string(ParentPath(newP))
	parentID, parentNode, ok := tx.findNodeByPath(parentPath)
	if !ok {
		return newError(ErrNotFound, "new parent path not found", "tx.rename", parentPath, -1, nil)
	}
	if !parentNode.IsStorage() {
		return newError(ErrInvalidArgument, "new parent is not a storage", "tx.rename", parentPath, -1, nil)
	}
	tx.topologyChanged = true

	oldParentID := node.ParentID
	node.Path = string(newP)
	node.Name = pathBase(string(newP))
	node.ParentID = parentID
	tx.nodes[id] = node
	if e, ok := tx.entries[id]; ok {
		e.Name = node.Name
		tx.entries[id] = e
	}

	oldPrefix := string(oldP)
	newPrefix := string(newP)
	oldLower := strings.ToLower(oldPrefix)
	for _, did := range tx.collectSubtree(id) {
		if did == id {
			continue
		}
		n := tx.nodes[did]
		pLower := strings.ToLower(n.Path)
		if pLower == oldLower || strings.HasPrefix(pLower, oldLower+"/") {
			suffix := n.Path[len(oldPrefix):]
			n.Path = newPrefix + suffix
			tx.nodes[did] = n
		}
	}

	tx.children[oldParentID] = filterNodeID(tx.children[oldParentID], []NodeID{id})
	tx.children[parentID] = filterNodeID(tx.children[parentID], []NodeID{id})
	tx.children[parentID] = append(tx.children[parentID], id)

	if p, ok := tx.nodes[oldParentID]; ok {
		p.ChildCount = len(tx.children[oldParentID])
		tx.nodes[oldParentID] = p
	}
	parentNode.ChildCount = len(tx.children[parentID])
	tx.nodes[parentID] = parentNode
	tx.rebuildOrder()
	return nil
}

func (tx *Tx) CreateStorage(storagePath string) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.create_storage", storagePath, -1, nil)
	}
	p, err := CanonicalPath(storagePath)
	if err != nil {
		return err
	}
	if p == "/" {
		return nil
	}
	if _, _, exists := tx.findNodeByPath(string(p)); exists {
		return newError(ErrConflict, "path already exists", "tx.create_storage", string(p), -1, nil)
	}
	parentPath := string(ParentPath(p))
	parentID, parentNode, ok := tx.findNodeByPath(parentPath)
	if !ok {
		return newError(ErrNotFound, "parent path not found", "tx.create_storage", parentPath, -1, nil)
	}
	if !parentNode.IsStorage() {
		return newError(ErrInvalidArgument, "parent is not a storage", "tx.create_storage", parentPath, -1, nil)
	}
	if tx.file.opt.MaxObjectCount > 0 && len(tx.nodes)+1 > tx.file.opt.MaxObjectCount {
		return newError(ErrLimitExceeded, "object count exceeded limit", "tx.create_storage", string(p), -1, nil)
	}
	tx.topologyChanged = true

	id := tx.allocateID()
	name := pathBase(string(p))
	n := Node{
		ID:       id,
		Type:     NodeStorage,
		Path:     string(p),
		Name:     name,
		ParentID: parentID,
	}
	tx.nodes[id] = n
	tx.entries[id] = dirEntry{
		ID:           uint32(id),
		Name:         name,
		ObjectType:   1,
		LeftSibling:  cfbNoStream,
		RightSibling: cfbNoStream,
		Child:        cfbNoStream,
		StartSector:  cfbEndOfChain,
		Size:         0,
	}
	tx.children[parentID] = append(tx.children[parentID], id)
	parentNode.ChildCount = len(tx.children[parentID])
	tx.nodes[parentID] = parentNode
	tx.rebuildOrder()
	return nil
}

func (tx *Tx) Commit(ctx context.Context, opt CommitOptions) (*CommitResult, error) {
	if tx == nil || tx.closed {
		return nil, newError(ErrTxClosed, "transaction is closed", "tx.commit", "", -1, nil)
	}
	defer tx.release()
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
	if opt.Strategy == Incremental {
		result, err := tx.commitIncremental(ctx, opt)
		if err == nil {
			return result, nil
		}
		if !IsCode(err, ErrUnsupported) {
			return nil, newError(ErrCommitFailed, "incremental commit failed", "tx.commit", "", -1, err)
		}
		// Fallback to full rewrite when incremental preconditions are not met.
	}

	snapshot, err := tx.serializeFullRewrite()
	if err != nil {
		return nil, newError(ErrCommitFailed, "serialize failed", "tx.commit", "", -1, err)
	}
	if err := writeBackendBytes(tx.file.wb, snapshot); err != nil {
		return nil, newError(ErrCommitFailed, "write backend failed", "tx.commit", "", -1, err)
	}
	tx.file.mu.Lock()
	if err := tx.file.reloadStateFromBytes(snapshot); err != nil {
		tx.file.mu.Unlock()
		return nil, newError(ErrCommitFailed, "reload committed state failed", "tx.commit", "", -1, err)
	}
	tx.file.mu.Unlock()

	if tx.file.wb != nil && opt.Sync {
		if err := tx.file.wb.Sync(); err != nil {
			return nil, newError(ErrCommitFailed, "sync failed", "tx.commit", "", -1, err)
		}
	}
	size := int64(len(snapshot))
	return &CommitResult{
		BytesWritten: size,
		NewSize:      size,
		BackendKind:  backendKind(tx.file.rb),
	}, nil
}

func (tx *Tx) Revert() error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "tx.revert", "", -1, nil)
	}
	tx.closed = true
	tx.release()
	return nil
}

func (tx *Tx) release() {
	if tx == nil || tx.file == nil {
		return
	}
	tx.file.mu.Lock()
	defer tx.file.mu.Unlock()
	if tx.file.activeTx == tx {
		tx.file.activeTx = nil
	}
}

func (tx *Tx) commitIncremental(ctx context.Context, opt CommitOptions) (*CommitResult, error) {
	if tx.file == nil || tx.file.wb == nil || tx.file.hdr == nil {
		return nil, newError(ErrUnsupported, "incremental commit requires writable parsed container", "tx.commit_incremental", "", -1, nil)
	}
	if tx.topologyChanged {
		return nil, newError(ErrUnsupported, "incremental commit requires unchanged topology", "tx.commit_incremental", "", -1, nil)
	}
	ids := tx.sortedTouchedStreamIDs()
	for _, id := range ids {
		data, ok := tx.streamData[id]
		if !ok {
			return nil, newError(ErrUnsupported, "incremental stream data is missing", "tx.commit_incremental", "", -1, nil)
		}
		if err := tx.file.validateIncrementalStreamWrite(id, data); err != nil {
			return nil, err
		}
	}

	var bytesWritten int64
	for _, id := range ids {
		if errors.Is(ctx.Err(), context.Canceled) {
			return nil, newError(ErrCommitFailed, "incremental commit canceled", "tx.commit_incremental", "", -1, ctx.Err())
		}
		if errors.Is(ctx.Err(), context.DeadlineExceeded) {
			return nil, newError(ErrCommitFailed, "incremental commit timed out", "tx.commit_incremental", "", -1, ctx.Err())
		}
		data := tx.streamData[id]
		if err := tx.file.writeStreamByIDIncremental(id, data); err != nil {
			return nil, err
		}
		bytesWritten += int64(len(data))
	}

	tx.file.mu.Lock()
	tx.file.nodes = cloneNodes(tx.nodes)
	tx.file.order = append([]NodeID(nil), tx.order...)
	tx.file.children = cloneChildren(tx.children)
	tx.file.entries = cloneEntries(tx.entries)
	if root, ok := tx.file.nodes[0]; ok {
		tx.file.root = root
	}
	if tx.file.streamData == nil {
		tx.file.streamData = map[NodeID][]byte{}
	}
	for _, id := range ids {
		tx.file.streamData[id] = append([]byte(nil), tx.streamData[id]...)
	}
	tx.file.mu.Unlock()

	if tx.file.wb != nil && opt.Sync {
		if err := tx.file.wb.Sync(); err != nil {
			return nil, newError(ErrCommitFailed, "sync failed", "tx.commit_incremental", "", -1, err)
		}
	}
	size := tx.file.rb.Size()
	return &CommitResult{
		BytesWritten: bytesWritten,
		NewSize:      size,
		BackendKind:  backendKind(tx.file.rb),
	}, nil
}

func (tx *Tx) sortedTouchedStreamIDs() []NodeID {
	ids := make([]NodeID, 0, len(tx.touchedStreams))
	for id := range tx.touchedStreams {
		ids = append(ids, id)
	}
	sort.Slice(ids, func(i, j int) bool { return ids[i] < ids[j] })
	return ids
}

func (f *File) validateIncrementalStreamWrite(id NodeID, data []byte) error {
	node, ok := f.nodes[id]
	if !ok {
		return newError(ErrUnsupported, "incremental stream not found in base file", "tx.commit_incremental", "", -1, nil)
	}
	if !node.IsStream() {
		return newError(ErrUnsupported, "incremental target is not a stream", "tx.commit_incremental", node.Path, -1, nil)
	}
	if int64(len(data)) != node.Size {
		return newError(ErrUnsupported, "incremental commit requires unchanged stream size", "tx.commit_incremental", node.Path, -1, nil)
	}
	entry, ok := f.entries[id]
	if !ok {
		return newError(ErrUnsupported, "incremental stream directory entry is missing", "tx.commit_incremental", node.Path, -1, nil)
	}
	if len(data) == 0 {
		return nil
	}
	if uint64(len(data)) < uint64(f.hdr.MiniStreamCutoff) {
		if len(f.miniFAT) == 0 || entry.StartSector == cfbEndOfChain {
			return newError(ErrUnsupported, "mini stream chain is unavailable for incremental write", "tx.commit_incremental", node.Path, -1, nil)
		}
		if _, ok := f.entries[0]; !ok {
			return newError(ErrUnsupported, "root mini stream entry is missing", "tx.commit_incremental", "/", -1, nil)
		}
		return nil
	}
	if len(f.fat) == 0 || entry.StartSector == cfbEndOfChain {
		return newError(ErrUnsupported, "fat chain is unavailable for incremental write", "tx.commit_incremental", node.Path, -1, nil)
	}
	return nil
}

func (f *File) writeStreamByIDIncremental(id NodeID, data []byte) error {
	node := f.nodes[id]
	entry := f.entries[id]
	if len(data) == 0 {
		return nil
	}
	if uint64(len(data)) < uint64(f.hdr.MiniStreamCutoff) {
		return f.writeMiniStreamByID(node.Path, entry.StartSector, data)
	}
	return f.writeRegularStreamByID(node.Path, entry.StartSector, data)
}

func (f *File) writeRegularStreamByID(streamPath string, startSector uint32, data []byte) error {
	sectorSize := int64(1 << f.hdr.SectorShift)
	chain, err := walkFATChain(f.fat, startSector, f.opt.MaxChainLength)
	if err != nil {
		return newError(ErrUnsupported, "failed to resolve stream fat chain for incremental write", "tx.commit_incremental", streamPath, -1, err)
	}
	expected := ceilDiv(len(data), int(sectorSize))
	if len(chain) < expected {
		return newError(ErrUnsupported, "stream fat chain is shorter than payload sectors", "tx.commit_incremental", streamPath, -1, nil)
	}
	offset := 0
	for i := 0; i < expected; i++ {
		sid := chain[i]
		chunk := int(sectorSize)
		if rem := len(data) - offset; rem < chunk {
			chunk = rem
		}
		off := sectorOffset(sid, sectorSize)
		if err := writeAtAll(f.wb, data[offset:offset+chunk], off); err != nil {
			return newError(ErrCommitFailed, "write regular stream sector failed", "tx.commit_incremental", streamPath, off, err)
		}
		offset += chunk
	}
	return nil
}

func (f *File) writeMiniStreamByID(streamPath string, startMiniSector uint32, data []byte) error {
	miniChain, err := walkFATChain(f.miniFAT, startMiniSector, f.opt.MaxChainLength)
	if err != nil {
		return newError(ErrUnsupported, "failed to resolve mini fat chain for incremental write", "tx.commit_incremental", streamPath, -1, err)
	}
	miniSectorSize := int64(1 << f.hdr.MiniSectorShift)
	expectedMini := ceilDiv(len(data), int(miniSectorSize))
	if len(miniChain) < expectedMini {
		return newError(ErrUnsupported, "mini fat chain is shorter than payload sectors", "tx.commit_incremental", streamPath, -1, nil)
	}
	root, ok := f.entries[0]
	if !ok || root.StartSector == cfbEndOfChain {
		return newError(ErrUnsupported, "root mini stream chain is unavailable", "tx.commit_incremental", "/", -1, nil)
	}
	sectorSize := int64(1 << f.hdr.SectorShift)
	rootChain, err := walkFATChain(f.fat, root.StartSector, f.opt.MaxChainLength)
	if err != nil {
		return newError(ErrUnsupported, "failed to resolve root mini stream chain", "tx.commit_incremental", "/", -1, err)
	}

	offset := 0
	for i := 0; i < expectedMini; i++ {
		chunk := int(miniSectorSize)
		if rem := len(data) - offset; rem < chunk {
			chunk = rem
		}
		miniOff := int64(miniChain[i]) * miniSectorSize
		rootSectorIndex := int(miniOff / sectorSize)
		if rootSectorIndex >= len(rootChain) {
			return newError(ErrUnsupported, "root mini stream chain is shorter than required", "tx.commit_incremental", streamPath, miniOff, nil)
		}
		inner := miniOff % sectorSize
		off := sectorOffset(rootChain[rootSectorIndex], sectorSize) + inner
		if err := writeAtAll(f.wb, data[offset:offset+chunk], off); err != nil {
			return newError(ErrCommitFailed, "write mini stream sector failed", "tx.commit_incremental", streamPath, off, err)
		}
		if len(f.miniData) >= int(miniOff)+chunk {
			copy(f.miniData[int(miniOff):int(miniOff)+chunk], data[offset:offset+chunk])
		}
		offset += chunk
	}
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
		rb:         rb,
		wb:         wb,
		opt:        opt,
		root:       root,
		nodes:      map[NodeID]Node{0: root},
		order:      []NodeID{0},
		children:   map[NodeID][]NodeID{},
		entries:    map[NodeID]dirEntry{},
		streamData: map[NodeID][]byte{},
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

func cloneNodes(in map[NodeID]Node) map[NodeID]Node {
	out := make(map[NodeID]Node, len(in))
	for k, v := range in {
		out[k] = v
	}
	return out
}

func cloneEntries(in map[NodeID]dirEntry) map[NodeID]dirEntry {
	out := make(map[NodeID]dirEntry, len(in))
	for k, v := range in {
		out[k] = v
	}
	return out
}

func cloneChildren(in map[NodeID][]NodeID) map[NodeID][]NodeID {
	out := make(map[NodeID][]NodeID, len(in))
	for k, v := range in {
		out[k] = append([]NodeID(nil), v...)
	}
	return out
}

func cloneStreamData(in map[NodeID][]byte) map[NodeID][]byte {
	out := make(map[NodeID][]byte, len(in))
	for k, v := range in {
		out[k] = append([]byte(nil), v...)
	}
	return out
}

func maxNodeID(nodes map[NodeID]Node) NodeID {
	var max NodeID
	for id := range nodes {
		if id > max {
			max = id
		}
	}
	return max
}

func (tx *Tx) allocateID() NodeID {
	id := tx.nextID
	tx.nextID++
	return id
}

func (tx *Tx) findNodeByPath(p string) (NodeID, Node, bool) {
	for _, id := range tx.order {
		n, ok := tx.nodes[id]
		if !ok {
			continue
		}
		if strings.EqualFold(n.Path, p) {
			return id, n, true
		}
	}
	return 0, Node{}, false
}

func (tx *Tx) collectSubtree(rootID NodeID) []NodeID {
	out := make([]NodeID, 0, 8)
	stack := []NodeID{rootID}
	seen := map[NodeID]struct{}{}
	for len(stack) > 0 {
		id := stack[len(stack)-1]
		stack = stack[:len(stack)-1]
		if _, ok := seen[id]; ok {
			continue
		}
		seen[id] = struct{}{}
		out = append(out, id)
		for _, child := range tx.children[id] {
			stack = append(stack, child)
		}
	}
	return out
}

func (tx *Tx) rebuildOrder() {
	order := make([]NodeID, 0, len(tx.nodes))
	order = append(order, 0)
	var walk func(parent NodeID)
	walk = func(parent NodeID) {
		children := append([]NodeID(nil), tx.children[parent]...)
		sort.Slice(children, func(i, j int) bool {
			ni := tx.nodes[children[i]]
			nj := tx.nodes[children[j]]
			if !strings.EqualFold(ni.Path, nj.Path) {
				return strings.ToLower(ni.Path) < strings.ToLower(nj.Path)
			}
			return children[i] < children[j]
		})
		tx.children[parent] = children
		for _, cid := range children {
			n, ok := tx.nodes[cid]
			if !ok {
				continue
			}
			order = append(order, cid)
			if n.IsStorage() {
				walk(cid)
			}
		}
	}
	walk(0)
	tx.order = order
}

func (tx *Tx) bytesWritten() int64 {
	var n int64
	for _, data := range tx.streamData {
		n += int64(len(data))
	}
	return n
}

func filterNodeID(in []NodeID, remove []NodeID) []NodeID {
	if len(in) == 0 || len(remove) == 0 {
		return in
	}
	rm := make(map[NodeID]struct{}, len(remove))
	for _, id := range remove {
		rm[id] = struct{}{}
	}
	out := in[:0]
	for _, id := range in {
		if _, ok := rm[id]; ok {
			continue
		}
		out = append(out, id)
	}
	return out
}

func pathBase(p string) string {
	if p == "/" {
		return "/"
	}
	_, b := path.Split(p)
	name, err := DecodeSegment(b)
	if err != nil || name == "" {
		return b
	}
	return name
}

func writeBackendBytes(wb storage.WriteBackend, data []byte) error {
	if wb == nil {
		return newError(ErrReadOnly, "write backend is nil", "backend.write_all", "", -1, nil)
	}
	if err := writeAtAll(wb, data, 0); err != nil {
		return err
	}
	if err := wb.Truncate(int64(len(data))); err != nil {
		return err
	}
	return nil
}

func writeAtAll(wb storage.WriteBackend, p []byte, off int64) error {
	written := 0
	for written < len(p) {
		n, err := wb.WriteAt(p[written:], off+int64(written))
		if err != nil {
			return err
		}
		if n <= 0 {
			return io.ErrShortWrite
		}
		written += n
	}
	return nil
}

func (f *File) reloadStateFromBytes(snapshot []byte) error {
	reloaded, err := OpenBytes(snapshot, f.opt)
	if err != nil {
		return err
	}
	defer reloaded.Close()

	f.root = reloaded.root
	f.hdr = reloaded.hdr
	f.fat = append([]uint32(nil), reloaded.fat...)
	f.miniFAT = append([]uint32(nil), reloaded.miniFAT...)
	f.miniData = append([]byte(nil), reloaded.miniData...)
	f.nodes = cloneNodes(reloaded.nodes)
	f.order = append([]NodeID(nil), reloaded.order...)
	f.children = cloneChildren(reloaded.children)
	f.entries = cloneEntries(reloaded.entries)
	f.streamData = map[NodeID][]byte{}
	f.report = reloaded.report
	return nil
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
