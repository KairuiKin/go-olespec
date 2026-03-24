package olecfb

import (
	"encoding/binary"
	"sort"
	"strings"
	"unicode/utf16"
)

const (
	serializerDefaultSectorSize = 512
	serializerMiniSectorSize = 64
)

type normalizedGraph struct {
	order    []NodeID
	nodes    map[NodeID]Node
	children map[NodeID][]NodeID
	oldByNew map[NodeID]NodeID
}

type streamLayout struct {
	entryID NodeID
	data    []byte
	size    int64
	isMini  bool
	start   uint32
}

func (tx *Tx) serializeFullRewrite() ([]byte, error) {
	majorVersion, sectorShift, sectorSize := tx.targetContainerGeometry()
	graph, err := tx.normalizeGraph()
	if err != nil {
		return nil, err
	}
	streams, err := tx.collectStreamLayouts(graph)
	if err != nil {
		return nil, err
	}

	miniData, miniFAT, streamStarts, err := buildMiniStreamLayouts(streams)
	if err != nil {
		return nil, err
	}

	dirEntries := tx.buildDirEntries(graph)
	for _, st := range streams {
		e := dirEntries[st.entryID]
		e.Size = st.size
		if st.isMini {
			e.StartSector = streamStarts[st.entryID]
		}
		dirEntries[st.entryID] = e
	}

	regularStreams := make([]streamLayout, 0, len(streams))
	for _, st := range streams {
		if st.isMini {
			continue
		}
		regularStreams = append(regularStreams, st)
	}

	dirBytes, err := encodeDirEntries(dirEntries, graph.order)
	if err != nil {
		return nil, err
	}
	regularSectors := countRegularStreamSectors(regularStreams, sectorSize)
	miniStreamSectors := ceilDiv(len(miniData), sectorSize)
	miniFATSectors := ceilDiv(len(miniFAT)*4, sectorSize)
	dirSectors := ceilDiv(len(dirBytes), sectorSize)
	if dirSectors == 0 {
		dirSectors = 1
	}

	nonFATSectors := regularSectors + miniStreamSectors + miniFATSectors + dirSectors
	fatSectors, difatSectors := solveFATAndDIFATSectorCount(
		nonFATSectors,
		sectorSize/4,
		cfbNumDifatEntries,
		sectorSize/4-1,
	)
	if fatSectors <= 0 {
		fatSectors = 1
	}
	totalSectors := nonFATSectors + fatSectors + difatSectors
	fatEntries := make([]uint32, totalSectors)
	for i := range fatEntries {
		fatEntries[i] = cfbFreeSector
	}
	sectors := make([][]byte, totalSectors)
	for i := range sectors {
		sectors[i] = make([]byte, sectorSize)
	}

	cursor := uint32(0)
	for i := range regularStreams {
		st := &regularStreams[i]
		secCount := ceilDiv(len(st.data), sectorSize)
		if secCount == 0 {
			st.start = cfbEndOfChain
			e := dirEntries[st.entryID]
			e.StartSector = cfbEndOfChain
			e.Size = 0
			dirEntries[st.entryID] = e
			continue
		}
		st.start = cursor
		e := dirEntries[st.entryID]
		e.StartSector = st.start
		e.Size = st.size
		dirEntries[st.entryID] = e
		writeChainData(sectors, cursor, sectorSize, st.data)
		markChainFAT(fatEntries, cursor, secCount)
		cursor += uint32(secCount)
	}

	var miniStreamStart uint32 = cfbEndOfChain
	if miniStreamSectors > 0 {
		miniStreamStart = cursor
		writeChainData(sectors, cursor, sectorSize, miniData)
		markChainFAT(fatEntries, cursor, miniStreamSectors)
		cursor += uint32(miniStreamSectors)
	}

	var miniFATStart uint32 = cfbEndOfChain
	if miniFATSectors > 0 {
		miniFATStart = cursor
		miniRaw := make([]byte, miniFATSectors*sectorSize)
		for i, v := range miniFAT {
			binary.LittleEndian.PutUint32(miniRaw[i*4:i*4+4], v)
		}
		writeChainData(sectors, cursor, sectorSize, miniRaw)
		markChainFAT(fatEntries, cursor, miniFATSectors)
		cursor += uint32(miniFATSectors)
	}

	dirStart := cursor
	dirBytes, err = encodeDirEntries(dirEntries, graph.order)
	if err != nil {
		return nil, err
	}
	dirRaw := make([]byte, dirSectors*sectorSize)
	copy(dirRaw, dirBytes)
	writeChainData(sectors, cursor, sectorSize, dirRaw)
	markChainFAT(fatEntries, cursor, dirSectors)
	cursor += uint32(dirSectors)

	fatSectorIDs := make([]uint32, fatSectors)
	for i := 0; i < fatSectors; i++ {
		fatSectorIDs[i] = cursor + uint32(i)
	}
	for _, sid := range fatSectorIDs {
		fatEntries[sid] = cfbFatSector
	}
	cursor += uint32(fatSectors)

	difatSectorIDs := make([]uint32, difatSectors)
	for i := 0; i < difatSectors; i++ {
		difatSectorIDs[i] = cursor + uint32(i)
	}
	for _, sid := range difatSectorIDs {
		fatEntries[sid] = cfbDifatSector
	}
	cursor += uint32(difatSectors)
	if int(cursor) != totalSectors {
		return nil, newError(ErrCommitFailed, "sector allocation mismatch", "serialize.full_rewrite", "", -1, nil)
	}

	for i, sid := range fatSectorIDs {
		buf := sectors[sid]
		for j := 0; j < sectorSize/4; j++ {
			idx := i*(sectorSize/4) + j
			v := cfbFreeSector
			if idx < len(fatEntries) {
				v = fatEntries[idx]
			}
			binary.LittleEndian.PutUint32(buf[j*4:j*4+4], v)
		}
	}
	writeDIFATSectors(sectors, difatSectorIDs, fatSectorIDs, sectorSize/4-1)

	root := dirEntries[0]
	root.StartSector = miniStreamStart
	root.Size = int64(len(miniData))
	dirEntries[0] = root

	dirBytes, err = encodeDirEntries(dirEntries, graph.order)
	if err != nil {
		return nil, err
	}
	for i := 0; i < dirSectors; i++ {
		sid := dirStart + uint32(i)
		copy(sectors[sid], make([]byte, sectorSize))
		start := i * sectorSize
		end := start + sectorSize
		if end > len(dirBytes) {
			end = len(dirBytes)
		}
		if start < len(dirBytes) {
			copy(sectors[sid], dirBytes[start:end])
		}
	}

	header := buildSerializedHeader(
		majorVersion,
		sectorShift,
		fatSectorIDs,
		difatSectorIDs,
		dirStart,
		dirSectors,
		miniFATStart,
		miniFATSectors,
	)

	out := make([]byte, sectorSize+len(sectors)*sectorSize)
	copy(out[:cfbHeaderSize], header)
	for i := range sectors {
		start := sectorSize + i*sectorSize
		copy(out[start:start+sectorSize], sectors[i])
	}
	return out, nil
}

func (tx *Tx) targetContainerGeometry() (uint16, uint16, int) {
	if tx != nil && tx.file != nil && tx.file.hdr != nil && tx.file.hdr.MajorVersion == cfbMajorVersion4 {
		return cfbMajorVersion4, cfbSectorShiftV4, 1 << cfbSectorShiftV4
	}
	return cfbMajorVersion3, cfbSectorShiftV3, serializerDefaultSectorSize
}

func (tx *Tx) normalizeGraph() (*normalizedGraph, error) {
	if len(tx.nodes) == 0 {
		return nil, newError(ErrCommitFailed, "empty node set", "serialize.normalize", "", -1, nil)
	}
	order := make([]NodeID, 0, len(tx.nodes))
	seen := make(map[NodeID]struct{}, len(tx.nodes))
	for _, id := range tx.order {
		if _, ok := tx.nodes[id]; !ok {
			continue
		}
		if _, ok := seen[id]; ok {
			continue
		}
		seen[id] = struct{}{}
		order = append(order, id)
	}

	leftover := make([]NodeID, 0)
	for id := range tx.nodes {
		if _, ok := seen[id]; ok {
			continue
		}
		leftover = append(leftover, id)
	}
	sort.Slice(leftover, func(i, j int) bool {
		ni := tx.nodes[leftover[i]]
		nj := tx.nodes[leftover[j]]
		if !strings.EqualFold(ni.Path, nj.Path) {
			return strings.ToLower(ni.Path) < strings.ToLower(nj.Path)
		}
		return leftover[i] < leftover[j]
	})
	order = append(order, leftover...)

	rootPos := -1
	for i, id := range order {
		n := tx.nodes[id]
		if n.Type == NodeRoot || n.Path == "/" {
			rootPos = i
			break
		}
	}
	if rootPos < 0 {
		return nil, newError(ErrCommitFailed, "root node not found", "serialize.normalize", "/", -1, nil)
	}
	if rootPos != 0 {
		rootID := order[rootPos]
		order = append([]NodeID{rootID}, append(order[:rootPos], order[rootPos+1:]...)...)
	}

	oldToNew := make(map[NodeID]NodeID, len(order))
	oldByNew := make(map[NodeID]NodeID, len(order))
	for i, oldID := range order {
		newID := NodeID(i)
		oldToNew[oldID] = newID
		oldByNew[newID] = oldID
	}

	nodes := make(map[NodeID]Node, len(order))
	for i, oldID := range order {
		newID := NodeID(i)
		n := tx.nodes[oldID]
		n.ID = newID
		if newID == 0 {
			n.ParentID = 0
			n.Type = NodeRoot
			n.Path = "/"
		} else if mapped, ok := oldToNew[n.ParentID]; ok {
			n.ParentID = mapped
		} else {
			n.ParentID = 0
		}
		nodes[newID] = n
	}

	children := make(map[NodeID][]NodeID, len(nodes))
	for oldParent, oldChildren := range tx.children {
		newParent, ok := oldToNew[oldParent]
		if !ok {
			continue
		}
		list := make([]NodeID, 0, len(oldChildren))
		dedup := map[NodeID]struct{}{}
		for _, oldChild := range oldChildren {
			newChild, ok := oldToNew[oldChild]
			if !ok {
				continue
			}
			if _, exists := dedup[newChild]; exists {
				continue
			}
			dedup[newChild] = struct{}{}
			list = append(list, newChild)
		}
		sort.Slice(list, func(i, j int) bool {
			ni := nodes[list[i]]
			nj := nodes[list[j]]
			if !strings.EqualFold(ni.Path, nj.Path) {
				return strings.ToLower(ni.Path) < strings.ToLower(nj.Path)
			}
			return list[i] < list[j]
		})
		children[newParent] = list
	}

	for id, n := range nodes {
		if n.IsStorage() {
			if _, ok := children[id]; !ok {
				children[id] = nil
			}
			n.ChildCount = len(children[id])
		} else {
			n.ChildCount = 0
		}
		nodes[id] = n
	}

	newOrder := make([]NodeID, len(order))
	for i := range newOrder {
		newOrder[i] = NodeID(i)
	}
	return &normalizedGraph{
		order:    newOrder,
		nodes:    nodes,
		children: children,
		oldByNew: oldByNew,
	}, nil
}

func (tx *Tx) collectStreamLayouts(graph *normalizedGraph) ([]streamLayout, error) {
	streams := make([]streamLayout, 0, len(graph.order))
	for _, id := range graph.order {
		n := graph.nodes[id]
		if !n.IsStream() {
			continue
		}
		oldID := graph.oldByNew[id]
		data, err := tx.loadStreamData(oldID, n.Size)
		if err != nil {
			return nil, err
		}
		st := streamLayout{
			entryID: id,
			data:    data,
			size:    int64(len(data)),
		}
		if st.size < cfbMiniStreamCutoff {
			st.isMini = true
		}
		streams = append(streams, st)
	}
	sort.Slice(streams, func(i, j int) bool {
		return streams[i].entryID < streams[j].entryID
	})
	return streams, nil
}

func buildMiniStreamLayouts(streams []streamLayout) ([]byte, []uint32, map[NodeID]uint32, error) {
	if len(streams) == 0 {
		return nil, nil, map[NodeID]uint32{}, nil
	}
	miniData := make([]byte, 0)
	miniFAT := make([]uint32, 0)
	startMap := make(map[NodeID]uint32)
	for i := range streams {
		st := streams[i]
		if !st.isMini {
			continue
		}
		if st.size == 0 {
			startMap[st.entryID] = cfbEndOfChain
			continue
		}
		if st.size < 0 {
			return nil, nil, nil, newError(ErrInvalidArgument, "negative stream size", "serialize.ministream", "", -1, nil)
		}
		secCount := ceilDiv(int(st.size), serializerMiniSectorSize)
		start := len(miniFAT)
		startMap[st.entryID] = uint32(start)

		padded := make([]byte, secCount*serializerMiniSectorSize)
		copy(padded, st.data)
		miniData = append(miniData, padded...)

		for j := 0; j < secCount; j++ {
			next := cfbEndOfChain
			if j+1 < secCount {
				next = uint32(start + j + 1)
			}
			miniFAT = append(miniFAT, next)
		}
	}
	return miniData, miniFAT, startMap, nil
}

func (tx *Tx) buildDirEntries(graph *normalizedGraph) map[NodeID]dirEntry {
	out := make(map[NodeID]dirEntry, len(graph.nodes))
	for _, id := range graph.order {
		n := graph.nodes[id]
		oldID := graph.oldByNew[id]
		e := dirEntry{
			ID:           uint32(id),
			Name:         n.Name,
			LeftSibling:  cfbNoStream,
			RightSibling: cfbNoStream,
			Child:        cfbNoStream,
			CLSID:        n.CLSID,
			StateBits:    n.StateBits,
			CreatedAt:    n.CreatedAt,
			ModifiedAt:   n.ModifiedAt,
			StartSector:  cfbEndOfChain,
			Size:         n.Size,
		}
		if old, ok := tx.entries[oldID]; ok {
			if e.Name == "" {
				e.Name = old.Name
			}
			e.CLSID = old.CLSID
			e.StateBits = old.StateBits
			e.CreatedAt = old.CreatedAt
			e.ModifiedAt = old.ModifiedAt
		}
		switch n.Type {
		case NodeRoot:
			e.ObjectType = 5
			e.Name = "Root Entry"
			e.StartSector = cfbEndOfChain
			e.Size = 0
		case NodeStorage:
			e.ObjectType = 1
			e.StartSector = cfbEndOfChain
			e.Size = 0
		case NodeStream:
			e.ObjectType = 2
		default:
			e.ObjectType = 0
			e.StartSector = cfbEndOfChain
			e.Size = 0
		}
		out[id] = e
	}

	for parentID, kids := range graph.children {
		if len(kids) == 0 {
			continue
		}
		parent := out[parentID]
		parent.Child = uint32(kids[0])
		out[parentID] = parent
		for i, kidID := range kids {
			e := out[kidID]
			e.LeftSibling = cfbNoStream
			if i+1 < len(kids) {
				e.RightSibling = uint32(kids[i+1])
			} else {
				e.RightSibling = cfbNoStream
			}
			out[kidID] = e
		}
	}
	return out
}

func (tx *Tx) loadStreamData(oldID NodeID, sizeHint int64) ([]byte, error) {
	if data, ok := tx.streamData[oldID]; ok {
		return append([]byte(nil), data...), nil
	}
	if sizeHint <= 0 {
		return nil, nil
	}
	return tx.file.readStreamByID(oldID, sizeHint)
}

func (f *File) readStreamByID(id NodeID, sizeHint int64) ([]byte, error) {
	node, ok := f.nodes[id]
	if !ok {
		return nil, newError(ErrNotFound, "stream node not found", "stream.read_by_id", "", -1, nil)
	}
	if !node.IsStream() {
		return nil, newError(ErrInvalidArgument, "node is not a stream", "stream.read_by_id", node.Path, -1, nil)
	}
	if data, ok := f.streamData[id]; ok {
		return append([]byte(nil), data...), nil
	}
	if sizeHint <= 0 && node.Size <= 0 {
		return nil, nil
	}
	if f.hdr == nil {
		return nil, newError(ErrBadHeader, "header is missing", "stream.read_by_id", node.Path, -1, nil)
	}

	entry, ok := f.entries[id]
	if !ok {
		return nil, newError(ErrDirCorrupt, "missing stream directory entry", "stream.read_by_id", node.Path, -1, nil)
	}
	size := node.Size
	if sizeHint >= 0 {
		size = sizeHint
	}
	if size == 0 {
		return nil, nil
	}
	if size < 0 {
		return nil, newError(ErrDirCorrupt, "negative stream size", "stream.read_by_id", node.Path, -1, nil)
	}

	if uint64(size) < uint64(f.hdr.MiniStreamCutoff) {
		if entry.StartSector == cfbEndOfChain {
			return nil, newError(ErrMiniStreamCorrupt, "mini stream start sector is missing", "stream.read_by_id", node.Path, -1, nil)
		}
		data, err := readMiniStreamData(f.miniData, f.miniFAT, f.hdr, entry.StartSector, size, f.opt.MaxChainLength)
		if err != nil {
			return nil, err
		}
		return data, nil
	}
	if entry.StartSector == cfbEndOfChain {
		return nil, newError(ErrBadFATChain, "stream has size but no sector chain", "stream.read_by_id", node.Path, -1, nil)
	}
	return readNormalStreamData(f.rb.ReadAt, f.rb.Size(), f.fat, f.hdr, entry.StartSector, size, f.opt.MaxChainLength)
}

func countRegularStreamSectors(streams []streamLayout, sectorSize int) int {
	total := 0
	for _, st := range streams {
		total += ceilDiv(len(st.data), sectorSize)
	}
	return total
}

func solveFATAndDIFATSectorCount(nonMetaSectors, entriesPerFatSector, headerDifatSlots, entriesPerDifatSector int) (int, int) {
	fat := 0
	difat := 0
	for {
		needDifat := difatSectorsForFAT(fat, headerDifatSlots, entriesPerDifatSector)
		needFat := ceilDiv(nonMetaSectors+fat+needDifat, entriesPerFatSector)
		if needFat == fat && needDifat == difat {
			return fat, difat
		}
		fat = needFat
		difat = needDifat
	}
}

func difatSectorsForFAT(fatCount, headerDifatSlots, entriesPerDifatSector int) int {
	if fatCount <= headerDifatSlots {
		return 0
	}
	return ceilDiv(fatCount-headerDifatSlots, entriesPerDifatSector)
}

func markChainFAT(fat []uint32, start uint32, sectors int) {
	if sectors <= 0 {
		return
	}
	for i := 0; i < sectors; i++ {
		sid := start + uint32(i)
		next := cfbEndOfChain
		if i+1 < sectors {
			next = sid + 1
		}
		fat[sid] = next
	}
}

func writeChainData(sectors [][]byte, start uint32, sectorSize int, data []byte) {
	if len(data) == 0 {
		return
	}
	for i := 0; i < ceilDiv(len(data), sectorSize); i++ {
		sid := start + uint32(i)
		off := i * sectorSize
		end := off + sectorSize
		if end > len(data) {
			end = len(data)
		}
		copy(sectors[sid], data[off:end])
	}
}

func writeDIFATSectors(sectors [][]byte, difatSectorIDs []uint32, fatSectorIDs []uint32, entriesPerDifatSector int) {
	if len(difatSectorIDs) == 0 {
		return
	}
	remain := fatSectorIDs
	if len(remain) > cfbNumDifatEntries {
		remain = remain[cfbNumDifatEntries:]
	} else {
		remain = nil
	}
	cursor := 0
	for i, sid := range difatSectorIDs {
		buf := sectors[sid]
		for j := 0; j < entriesPerDifatSector; j++ {
			v := cfbFreeSector
			if cursor < len(remain) {
				v = remain[cursor]
				cursor++
			}
			binary.LittleEndian.PutUint32(buf[j*4:j*4+4], v)
		}
		next := cfbEndOfChain
		if i+1 < len(difatSectorIDs) {
			next = difatSectorIDs[i+1]
		}
		binary.LittleEndian.PutUint32(buf[entriesPerDifatSector*4:entriesPerDifatSector*4+4], next)
	}
}

func encodeDirEntries(entries map[NodeID]dirEntry, order []NodeID) ([]byte, error) {
	out := make([]byte, len(order)*cfbDirEntrySize)
	for i, id := range order {
		e, ok := entries[id]
		if !ok {
			e = dirEntry{
				ID:           uint32(id),
				ObjectType:   0,
				LeftSibling:  cfbNoStream,
				RightSibling: cfbNoStream,
				Child:        cfbNoStream,
				StartSector:  cfbEndOfChain,
			}
		}
		e.ID = uint32(i)
		if err := writeDirEntryRaw(out[i*cfbDirEntrySize:(i+1)*cfbDirEntrySize], e); err != nil {
			return nil, err
		}
	}
	return out, nil
}

func writeDirEntryRaw(dst []byte, e dirEntry) error {
	for i := range dst {
		dst[i] = 0
	}
	nameBytes, nameLen, err := encodeDirNameUTF16(e.Name)
	if err != nil {
		return err
	}
	copy(dst[0:64], nameBytes)
	binary.LittleEndian.PutUint16(dst[64:66], nameLen)
	dst[66] = e.ObjectType
	dst[67] = 1 // black
	binary.LittleEndian.PutUint32(dst[68:72], e.LeftSibling)
	binary.LittleEndian.PutUint32(dst[72:76], e.RightSibling)
	binary.LittleEndian.PutUint32(dst[76:80], e.Child)
	copy(dst[80:96], e.CLSID[:])
	binary.LittleEndian.PutUint32(dst[96:100], e.StateBits)
	binary.LittleEndian.PutUint64(dst[100:108], uint64(e.CreatedAt))
	binary.LittleEndian.PutUint64(dst[108:116], uint64(e.ModifiedAt))
	binary.LittleEndian.PutUint32(dst[116:120], e.StartSector)
	binary.LittleEndian.PutUint64(dst[120:128], uint64(e.Size))
	return nil
}

func encodeDirNameUTF16(name string) ([]byte, uint16, error) {
	if name == "" {
		return make([]byte, 64), 0, nil
	}
	u16 := utf16.Encode([]rune(name))
	if len(u16) > 31 {
		return nil, 0, newError(ErrInvalidArgument, "directory name exceeds 31 UTF-16 code units", "serialize.dir_name", name, -1, nil)
	}
	raw := make([]byte, 64)
	for i, v := range u16 {
		binary.LittleEndian.PutUint16(raw[i*2:i*2+2], v)
	}
	length := uint16(len(u16)*2 + 2)
	return raw, length, nil
}

func buildSerializedHeader(
	majorVersion uint16,
	sectorShift uint16,
	fatSectorIDs []uint32,
	difatSectorIDs []uint32,
	dirStart uint32,
	dirSectors int,
	miniFATStart uint32,
	miniFATSectors int,
) []byte {
	buf := make([]byte, cfbHeaderSize)
	copy(buf[0:8], cfbSignature[:])
	binary.LittleEndian.PutUint16(buf[24:26], 0x003E)
	binary.LittleEndian.PutUint16(buf[26:28], majorVersion)
	binary.LittleEndian.PutUint16(buf[28:30], cfbByteOrder)
	binary.LittleEndian.PutUint16(buf[30:32], sectorShift)
	binary.LittleEndian.PutUint16(buf[32:34], cfbMiniSectorShift)
	if majorVersion == cfbMajorVersion4 {
		binary.LittleEndian.PutUint32(buf[40:44], uint32(dirSectors))
	} else {
		binary.LittleEndian.PutUint32(buf[40:44], 0) // v3 must be zero
	}
	binary.LittleEndian.PutUint32(buf[44:48], uint32(len(fatSectorIDs)))
	binary.LittleEndian.PutUint32(buf[48:52], dirStart)
	binary.LittleEndian.PutUint32(buf[52:56], 0)
	binary.LittleEndian.PutUint32(buf[56:60], cfbMiniStreamCutoff)
	binary.LittleEndian.PutUint32(buf[60:64], miniFATStart)
	binary.LittleEndian.PutUint32(buf[64:68], uint32(miniFATSectors))
	firstDifat := uint32(cfbEndOfChain)
	if len(difatSectorIDs) > 0 {
		firstDifat = difatSectorIDs[0]
	}
	binary.LittleEndian.PutUint32(buf[68:72], firstDifat)
	binary.LittleEndian.PutUint32(buf[72:76], uint32(len(difatSectorIDs)))
	for i := 0; i < cfbNumDifatEntries; i++ {
		v := cfbFreeSector
		if i < len(fatSectorIDs) {
			v = fatSectorIDs[i]
		}
		binary.LittleEndian.PutUint32(buf[76+i*4:80+i*4], v)
	}
	return buf
}

func ceilDiv(n, d int) int {
	if d <= 0 {
		return 0
	}
	if n <= 0 {
		return 0
	}
	return (n + d - 1) / d
}
