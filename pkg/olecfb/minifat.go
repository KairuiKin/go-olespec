package olecfb

import (
	"encoding/binary"
	"fmt"
)

func loadMiniData(
	readAt func([]byte, int64) (int, error),
	fileSize int64,
	hdr *cfbHeader,
	fat []uint32,
	entries map[NodeID]dirEntry,
	maxChainLength int,
) ([]uint32, []byte, error) {
	miniFAT, err := loadMiniFAT(readAt, fileSize, hdr, fat, maxChainLength)
	if err != nil {
		return nil, nil, err
	}
	if len(miniFAT) == 0 {
		return nil, nil, nil
	}

	rootEntry, ok := entries[0]
	if !ok {
		return nil, nil, newError(ErrDirCorrupt, "missing root entry for mini stream", "parse.ministream", "/", -1, nil)
	}
	if rootEntry.StartSector == cfbEndOfChain || rootEntry.Size == 0 {
		return miniFAT, nil, nil
	}

	miniData, err := readNormalStreamData(
		readAt,
		fileSize,
		fat,
		hdr,
		rootEntry.StartSector,
		rootEntry.Size,
		maxChainLength,
	)
	if err != nil {
		return nil, nil, newError(ErrMiniStreamCorrupt, "failed to read root mini stream", "parse.ministream", "/", -1, err)
	}
	return miniFAT, miniData, nil
}

func loadMiniFAT(
	readAt func([]byte, int64) (int, error),
	fileSize int64,
	hdr *cfbHeader,
	fat []uint32,
	maxChainLength int,
) ([]uint32, error) {
	if hdr == nil {
		return nil, newError(ErrBadHeader, "header is nil", "parse.minifat", "", -1, nil)
	}
	if hdr.NumMiniFATSectors == 0 || hdr.FirstMiniFAT == cfbEndOfChain {
		return nil, nil
	}
	if len(fat) == 0 {
		return nil, newError(ErrBadFATChain, "fat table is empty", "parse.minifat", "", -1, nil)
	}

	chain, err := walkFATChain(fat, hdr.FirstMiniFAT, maxChainLength)
	if err != nil {
		return nil, newError(ErrBadFATChain, "failed to resolve mini fat chain", "parse.minifat", "", -1, err)
	}
	if uint32(len(chain)) < hdr.NumMiniFATSectors {
		return nil, newError(
			ErrBadFATChain,
			fmt.Sprintf("mini fat chain too short: have %d need %d", len(chain), hdr.NumMiniFATSectors),
			"parse.minifat",
			"",
			-1,
			nil,
		)
	}

	sectorSize := int64(1 << hdr.SectorShift)
	entriesPerSector := int(sectorSize / 4)
	out := make([]uint32, 0, int(hdr.NumMiniFATSectors)*entriesPerSector)
	for i := uint32(0); i < hdr.NumMiniFATSectors; i++ {
		sid := chain[i]
		off := sectorOffset(sid, sectorSize)
		if off < 0 || off+sectorSize > fileSize {
			return nil, newError(ErrOutOfBounds, "mini fat sector out of bounds", "parse.minifat", "", off, nil)
		}
		buf := make([]byte, sectorSize)
		if err := readFullAt(readAt, buf, off); err != nil {
			return nil, newError(ErrMiniStreamCorrupt, "failed to read mini fat sector", "parse.minifat", "", off, err)
		}
		for j := 0; j < entriesPerSector; j++ {
			start := j * 4
			out = append(out, binary.LittleEndian.Uint32(buf[start:start+4]))
		}
	}
	return out, nil
}

func readMiniStreamData(
	miniStream []byte,
	miniFAT []uint32,
	hdr *cfbHeader,
	startMiniSector uint32,
	size int64,
	maxChainLength int,
) ([]byte, error) {
	if hdr == nil {
		return nil, newError(ErrBadHeader, "header is nil", "stream.read.mini", "", -1, nil)
	}
	if size == 0 {
		return nil, nil
	}
	if len(miniFAT) == 0 {
		return nil, newError(ErrMiniStreamCorrupt, "mini fat table is empty", "stream.read.mini", "", -1, nil)
	}
	if len(miniStream) == 0 {
		return nil, newError(ErrMiniStreamCorrupt, "mini stream is empty", "stream.read.mini", "", -1, nil)
	}

	chain, err := walkFATChain(miniFAT, startMiniSector, maxChainLength)
	if err != nil {
		return nil, newError(ErrMiniStreamCorrupt, "failed to resolve mini stream chain", "stream.read.mini", "", -1, err)
	}
	if len(chain) == 0 {
		return nil, newError(ErrMiniStreamCorrupt, "mini stream chain is empty", "stream.read.mini", "", -1, nil)
	}

	miniSectorSize := int64(1 << hdr.MiniSectorShift)
	expectedSectors := int((size + miniSectorSize - 1) / miniSectorSize)
	if len(chain) < expectedSectors {
		return nil, newError(
			ErrMiniStreamCorrupt,
			fmt.Sprintf("mini chain too short: have %d need %d", len(chain), expectedSectors),
			"stream.read.mini",
			"",
			-1,
			nil,
		)
	}

	out := make([]byte, size)
	written := int64(0)
	for _, sid := range chain {
		if written >= size {
			break
		}
		off := int64(sid) * miniSectorSize
		if off < 0 || off+miniSectorSize > int64(len(miniStream)) {
			return nil, newError(ErrOutOfBounds, "mini stream sector out of bounds", "stream.read.mini", "", off, nil)
		}
		toRead := miniSectorSize
		if rem := size - written; rem < toRead {
			toRead = rem
		}
		copy(out[written:written+toRead], miniStream[off:off+toRead])
		written += toRead
	}
	return out, nil
}
