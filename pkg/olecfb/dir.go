package olecfb

import (
	"encoding/binary"
	"fmt"
	"unicode/utf16"
)

const (
	cfbDirEntrySize = 128
	cfbNoStream     = 0xFFFFFFFF
)

type dirEntry struct {
	ID           uint32
	Name         string
	ObjectType   uint8
	LeftSibling  uint32
	RightSibling uint32
	Child        uint32
	CLSID        [16]byte
	StateBits    uint32
	CreatedAt    int64
	ModifiedAt   int64
	StartSector  uint32
	Size         int64
}

func parseRootNode(readAt func([]byte, int64) (int, error), size int64, hdr *cfbHeader) (Node, error) {
	root := Node{
		ID:   0,
		Type: NodeRoot,
		Path: "/",
		Name: "Root Entry",
	}
	if hdr == nil {
		return root, newError(ErrBadHeader, "header is nil", "parse.root", "/", -1, nil)
	}
	if hdr.FirstDirectory == cfbEndOfChain {
		// Empty container placeholder.
		return root, nil
	}
	sectorSize := int64(1 << hdr.SectorShift)
	offset := int64(hdr.FirstDirectory+1) * sectorSize
	if offset < 0 || offset+cfbDirEntrySize > size {
		return root, newError(ErrOutOfBounds, "directory entry offset is out of file bounds", "parse.root", "/", offset, nil)
	}
	buf := make([]byte, cfbDirEntrySize)
	if err := readFullAt(readAt, buf, offset); err != nil {
		return root, newError(ErrDirCorrupt, "failed to read root directory entry", "parse.root", "/", offset, err)
	}

	name, err := parseDirEntryName(buf[0:64], binary.LittleEndian.Uint16(buf[64:66]))
	if err != nil {
		return root, newError(ErrDirCorrupt, err.Error(), "parse.root", "/", offset, err)
	}
	objType := buf[66]
	if objType != 5 {
		return root, newError(ErrDirCorrupt, fmt.Sprintf("unexpected root object type: %d", objType), "parse.root", "/", offset+66, nil)
	}

	copy(root.CLSID[:], buf[80:96])
	root.StateBits = binary.LittleEndian.Uint32(buf[96:100])
	root.CreatedAt = int64(binary.LittleEndian.Uint64(buf[100:108]))
	root.ModifiedAt = int64(binary.LittleEndian.Uint64(buf[108:116]))
	root.Size = int64(binary.LittleEndian.Uint64(buf[120:128]))
	if name != "" {
		root.Name = name
	}
	return root, nil
}

func parseDirectoryEntries(readAt func([]byte, int64) (int, error), size int64, hdr *cfbHeader, fat []uint32, maxChainLength int) ([]dirEntry, error) {
	if hdr == nil {
		return nil, newError(ErrBadHeader, "header is nil", "parse.dir.entries", "", -1, nil)
	}
	if hdr.FirstDirectory == cfbEndOfChain {
		return nil, nil
	}

	sectorSize := int64(1 << hdr.SectorShift)
	var dirData []byte

	if len(fat) > 0 {
		chain, err := walkFATChain(fat, hdr.FirstDirectory, maxChainLength)
		if err != nil {
			return nil, newError(ErrBadFATChain, "failed to resolve directory chain", "parse.dir.entries", "", -1, err)
		}
		if len(chain) == 0 {
			return nil, nil
		}
		dirData = make([]byte, 0, len(chain)*int(sectorSize))
		for _, sid := range chain {
			off := sectorOffset(sid, sectorSize)
			if off < 0 || off+sectorSize > size {
				return nil, newError(ErrOutOfBounds, "directory sector out of bounds", "parse.dir.entries", "", off, nil)
			}
			buf := make([]byte, sectorSize)
			if err := readFullAt(readAt, buf, off); err != nil {
				return nil, newError(ErrDirCorrupt, "failed to read directory sector", "parse.dir.entries", "", off, err)
			}
			dirData = append(dirData, buf...)
		}
	} else {
		// Fallback path for minimal files without FAT data.
		off := sectorOffset(hdr.FirstDirectory, sectorSize)
		if off < 0 || off+sectorSize > size {
			return nil, newError(ErrOutOfBounds, "directory sector out of bounds", "parse.dir.entries", "", off, nil)
		}
		dirData = make([]byte, sectorSize)
		if err := readFullAt(readAt, dirData, off); err != nil {
			return nil, newError(ErrDirCorrupt, "failed to read directory sector", "parse.dir.entries", "", off, err)
		}
	}

	if len(dirData)%cfbDirEntrySize != 0 {
		return nil, newError(ErrDirCorrupt, "directory stream size is not aligned to entry size", "parse.dir.entries", "", -1, nil)
	}

	count := len(dirData) / cfbDirEntrySize
	entries := make([]dirEntry, 0, count)
	for i := 0; i < count; i++ {
		start := i * cfbDirEntrySize
		entry, err := parseDirectoryEntry(uint32(i), dirData[start:start+cfbDirEntrySize])
		if err != nil {
			return nil, newError(ErrDirCorrupt, err.Error(), "parse.dir.entries", "", int64(start), err)
		}
		entries = append(entries, entry)
	}
	return entries, nil
}

func parseDirectoryEntry(id uint32, raw []byte) (dirEntry, error) {
	if len(raw) != cfbDirEntrySize {
		return dirEntry{}, fmt.Errorf("invalid entry size: %d", len(raw))
	}
	name, err := parseDirEntryName(raw[0:64], binary.LittleEndian.Uint16(raw[64:66]))
	if err != nil {
		return dirEntry{}, err
	}
	e := dirEntry{
		ID:           id,
		Name:         name,
		ObjectType:   raw[66],
		LeftSibling:  binary.LittleEndian.Uint32(raw[68:72]),
		RightSibling: binary.LittleEndian.Uint32(raw[72:76]),
		Child:        binary.LittleEndian.Uint32(raw[76:80]),
		StateBits:    binary.LittleEndian.Uint32(raw[96:100]),
		CreatedAt:    int64(binary.LittleEndian.Uint64(raw[100:108])),
		ModifiedAt:   int64(binary.LittleEndian.Uint64(raw[108:116])),
		StartSector:  binary.LittleEndian.Uint32(raw[116:120]),
		Size:         int64(binary.LittleEndian.Uint64(raw[120:128])),
	}
	copy(e.CLSID[:], raw[80:96])
	return e, nil
}

func buildNodesFromEntries(entries []dirEntry, maxRecursion, maxObjectCount int) (Node, map[NodeID]Node, []NodeID, error) {
	root := Node{
		ID:   0,
		Type: NodeRoot,
		Path: "/",
		Name: "Root Entry",
	}
	if len(entries) == 0 {
		nodes := map[NodeID]Node{0: root}
		return root, nodes, []NodeID{0}, nil
	}
	if entries[0].ObjectType != 5 {
		return root, nil, nil, newError(ErrDirCorrupt, fmt.Sprintf("unexpected root object type: %d", entries[0].ObjectType), "parse.dir.tree", "/", -1, nil)
	}
	if entries[0].Name != "" {
		root.Name = entries[0].Name
	}
	root.CLSID = entries[0].CLSID
	root.StateBits = entries[0].StateBits
	root.CreatedAt = entries[0].CreatedAt
	root.ModifiedAt = entries[0].ModifiedAt
	root.Size = entries[0].Size

	nodes := map[NodeID]Node{0: root}
	order := []NodeID{0}
	entryUsed := map[uint32]struct{}{0: {}}

	var buildChildren func(parentID NodeID, parentPath Path, startID uint32, depth int) error
	buildChildren = func(parentID NodeID, parentPath Path, startID uint32, depth int) error {
		if startID == cfbNoStream {
			return nil
		}
		if maxRecursion > 0 && depth > maxRecursion {
			return newError(ErrDepthExceeded, "directory recursion depth exceeded", "parse.dir.tree", string(parentPath), -1, nil)
		}
		localSeen := map[uint32]struct{}{}
		var visitTree func(id uint32) error
		visitTree = func(id uint32) error {
			if id == cfbNoStream {
				return nil
			}
			if id >= uint32(len(entries)) {
				return newError(ErrOutOfBounds, fmt.Sprintf("directory id out of bounds: %d", id), "parse.dir.tree", string(parentPath), -1, nil)
			}
			if _, ok := localSeen[id]; ok {
				return newError(ErrCycleDetected, "cycle detected in directory siblings", "parse.dir.tree", string(parentPath), -1, nil)
			}
			localSeen[id] = struct{}{}
			e := entries[id]

			if err := visitTree(e.LeftSibling); err != nil {
				return err
			}

			if e.ObjectType == 1 || e.ObjectType == 2 {
				if _, ok := entryUsed[id]; ok {
					return newError(ErrCycleDetected, "directory entry reused in tree", "parse.dir.tree", string(parentPath), -1, nil)
				}
				entryUsed[id] = struct{}{}

				nodePath, err := JoinPath(parentPath, e.Name)
				if err != nil {
					return newError(ErrDirCorrupt, "failed to build node path", "parse.dir.tree", string(parentPath), -1, err)
				}
				node := Node{
					ID:         NodeID(e.ID),
					Path:       string(nodePath),
					Name:       e.Name,
					ParentID:   parentID,
					Size:       e.Size,
					CLSID:      e.CLSID,
					StateBits:  e.StateBits,
					CreatedAt:  e.CreatedAt,
					ModifiedAt: e.ModifiedAt,
				}
				if e.ObjectType == 1 {
					node.Type = NodeStorage
				} else {
					node.Type = NodeStream
				}
				nodes[node.ID] = node
				order = append(order, node.ID)
				if maxObjectCount > 0 && len(order) > maxObjectCount {
					return newError(ErrLimitExceeded, "object count exceeded limit", "parse.dir.tree", string(parentPath), -1, nil)
				}

				parent := nodes[parentID]
				parent.ChildCount++
				nodes[parentID] = parent

				if e.ObjectType == 1 && e.Child != cfbNoStream {
					if err := buildChildren(node.ID, nodePath, e.Child, depth+1); err != nil {
						return err
					}
				}
			}

			return visitTree(e.RightSibling)
		}
		return visitTree(startID)
	}

	if err := buildChildren(0, "/", entries[0].Child, 1); err != nil {
		return root, nil, nil, err
	}
	root = nodes[0]
	return root, nodes, order, nil
}

func parseDirEntryName(raw []byte, length uint16) (string, error) {
	if length == 0 {
		return "", nil
	}
	if int(length) > len(raw) || length < 2 || length%2 != 0 {
		return "", fmt.Errorf("invalid directory name length: %d", length)
	}
	// Name length includes terminating NUL (2 bytes).
	nameBytes := raw[:length-2]
	u16 := make([]uint16, len(nameBytes)/2)
	for i := range u16 {
		start := i * 2
		u16[i] = binary.LittleEndian.Uint16(nameBytes[start : start+2])
	}
	return string(utf16.Decode(u16)), nil
}
