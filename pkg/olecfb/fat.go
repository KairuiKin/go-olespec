package olecfb

import (
	"encoding/binary"
	"fmt"
)

func loadFAT(readAt func([]byte, int64) (int, error), size int64, hdr *cfbHeader) ([]uint32, error) {
	if hdr == nil {
		return nil, newError(ErrBadHeader, "header is nil", "parse.fat", "", -1, nil)
	}
	if hdr.NumFATSectors == 0 {
		return nil, nil
	}

	fatSectorIDs := make([]uint32, 0, hdr.NumFATSectors)
	for _, id := range hdr.DIFAT {
		if id == cfbFreeSector {
			continue
		}
		fatSectorIDs = append(fatSectorIDs, id)
		if uint32(len(fatSectorIDs)) == hdr.NumFATSectors {
			break
		}
	}

	if uint32(len(fatSectorIDs)) < hdr.NumFATSectors {
		if hdr.FirstDIFAT != cfbEndOfChain || hdr.NumDIFATSectors > 0 {
			return nil, newError(ErrUnsupported, "extended DIFAT chain is not implemented yet", "parse.fat", "", -1, nil)
		}
		return nil, newError(ErrBadHeader, "not enough FAT sector ids in DIFAT header", "parse.fat", "", -1, nil)
	}

	sectorSize := int64(1 << hdr.SectorShift)
	entriesPerSector := int(sectorSize / 4)
	fat := make([]uint32, 0, int(hdr.NumFATSectors)*entriesPerSector)

	for _, fatSectorID := range fatSectorIDs {
		offset := sectorOffset(fatSectorID, sectorSize)
		if offset < 0 || offset+sectorSize > size {
			return nil, newError(ErrOutOfBounds, fmt.Sprintf("fat sector %d out of bounds", fatSectorID), "parse.fat", "", offset, nil)
		}
		buf := make([]byte, sectorSize)
		if err := readFullAt(readAt, buf, offset); err != nil {
			return nil, newError(ErrBadFATChain, "failed to read FAT sector", "parse.fat", "", offset, err)
		}
		for i := 0; i < entriesPerSector; i++ {
			start := i * 4
			fat = append(fat, binary.LittleEndian.Uint32(buf[start:start+4]))
		}
	}
	return fat, nil
}

func sectorOffset(sectorID uint32, sectorSize int64) int64 {
	return int64(sectorID+1) * sectorSize
}

func walkFATChain(fat []uint32, start uint32, maxChainLength int) ([]uint32, error) {
	if start == cfbEndOfChain {
		return nil, nil
	}
	if len(fat) == 0 {
		return nil, newError(ErrBadFATChain, "fat table is empty", "fat.walk", "", -1, nil)
	}

	seen := make(map[uint32]struct{})
	chain := make([]uint32, 0, 8)
	current := start

	for {
		if maxChainLength > 0 && len(chain) >= maxChainLength {
			return nil, newError(ErrLimitExceeded, "fat chain length exceeded limit", "fat.walk", "", -1, nil)
		}
		if current >= uint32(len(fat)) {
			return nil, newError(ErrOutOfBounds, fmt.Sprintf("fat entry out of bounds: %d", current), "fat.walk", "", -1, nil)
		}
		if _, ok := seen[current]; ok {
			return nil, newError(ErrCycleDetected, "cycle detected in fat chain", "fat.walk", "", -1, nil)
		}
		seen[current] = struct{}{}
		chain = append(chain, current)

		next := fat[current]
		switch next {
		case cfbEndOfChain:
			return chain, nil
		case cfbFreeSector:
			return nil, newError(ErrBadFATChain, "fat chain points to free sector", "fat.walk", "", -1, nil)
		case cfbFatSector, cfbDifatSector:
			return nil, newError(ErrBadFATChain, "fat chain points to reserved sector type", "fat.walk", "", -1, nil)
		}
		current = next
	}
}
