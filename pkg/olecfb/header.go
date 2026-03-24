package olecfb

import (
	"encoding/binary"
	"fmt"
)

const (
	cfbHeaderSize              = 512
	cfbMiniStreamCutoff        = 4096
	cfbByteOrder               = 0xFFFE
	cfbMiniSectorShift         = 0x0006
	cfbMajorVersion3           = 0x0003
	cfbMajorVersion4           = 0x0004
	cfbSectorShiftV3           = 0x0009
	cfbSectorShiftV4           = 0x000C
	cfbNumDifatEntries         = 109
	cfbFreeSector       uint32 = 0xFFFFFFFF
	cfbEndOfChain       uint32 = 0xFFFFFFFE
	cfbFatSector        uint32 = 0xFFFFFFFD
	cfbDifatSector      uint32 = 0xFFFFFFFC
)

var cfbSignature = [8]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

type cfbHeader struct {
	Signature           [8]byte
	CLSID               [16]byte
	MinorVersion        uint16
	MajorVersion        uint16
	ByteOrder           uint16
	SectorShift         uint16
	MiniSectorShift     uint16
	Reserved            [6]byte
	NumDirectorySectors uint32
	NumFATSectors       uint32
	FirstDirectory      uint32
	TransactionSig      uint32
	MiniStreamCutoff    uint32
	FirstMiniFAT        uint32
	NumMiniFATSectors   uint32
	FirstDIFAT          uint32
	NumDIFATSectors     uint32
	DIFAT               [cfbNumDifatEntries]uint32
}

func parseHeader(readAt func([]byte, int64) (int, error), size int64) (*cfbHeader, error) {
	if size < cfbHeaderSize {
		return nil, newError(ErrBadHeader, "file is smaller than CFB header", "parse.header", "", -1, nil)
	}
	buf := make([]byte, cfbHeaderSize)
	if err := readFullAt(readAt, buf, 0); err != nil {
		return nil, newError(ErrBadHeader, "failed to read CFB header", "parse.header", "", 0, err)
	}

	var h cfbHeader
	copy(h.Signature[:], buf[0:8])
	copy(h.CLSID[:], buf[8:24])
	h.MinorVersion = binary.LittleEndian.Uint16(buf[24:26])
	h.MajorVersion = binary.LittleEndian.Uint16(buf[26:28])
	h.ByteOrder = binary.LittleEndian.Uint16(buf[28:30])
	h.SectorShift = binary.LittleEndian.Uint16(buf[30:32])
	h.MiniSectorShift = binary.LittleEndian.Uint16(buf[32:34])
	copy(h.Reserved[:], buf[34:40])
	h.NumDirectorySectors = binary.LittleEndian.Uint32(buf[40:44])
	h.NumFATSectors = binary.LittleEndian.Uint32(buf[44:48])
	h.FirstDirectory = binary.LittleEndian.Uint32(buf[48:52])
	h.TransactionSig = binary.LittleEndian.Uint32(buf[52:56])
	h.MiniStreamCutoff = binary.LittleEndian.Uint32(buf[56:60])
	h.FirstMiniFAT = binary.LittleEndian.Uint32(buf[60:64])
	h.NumMiniFATSectors = binary.LittleEndian.Uint32(buf[64:68])
	h.FirstDIFAT = binary.LittleEndian.Uint32(buf[68:72])
	h.NumDIFATSectors = binary.LittleEndian.Uint32(buf[72:76])
	for i := 0; i < cfbNumDifatEntries; i++ {
		start := 76 + i*4
		h.DIFAT[i] = binary.LittleEndian.Uint32(buf[start : start+4])
	}

	if err := validateHeader(&h); err != nil {
		return nil, err
	}
	return &h, nil
}

func validateHeader(h *cfbHeader) error {
	if h == nil {
		return newError(ErrBadHeader, "header is nil", "parse.header.validate", "", -1, nil)
	}
	if h.Signature != cfbSignature {
		return newError(ErrBadHeader, "invalid CFB signature", "parse.header.validate", "", 0, nil)
	}
	if h.ByteOrder != cfbByteOrder {
		return newError(ErrBadHeader, fmt.Sprintf("unexpected byte order: 0x%04X", h.ByteOrder), "parse.header.validate", "", 28, nil)
	}
	if h.MiniSectorShift != cfbMiniSectorShift {
		return newError(ErrBadHeader, fmt.Sprintf("unexpected mini sector shift: %d", h.MiniSectorShift), "parse.header.validate", "", 32, nil)
	}
	if h.MiniStreamCutoff != cfbMiniStreamCutoff {
		return newError(ErrBadHeader, fmt.Sprintf("unexpected mini stream cutoff: %d", h.MiniStreamCutoff), "parse.header.validate", "", 56, nil)
	}
	for i, b := range h.Reserved {
		if b != 0 {
			return newError(ErrBadHeader, "reserved bytes must be zero", "parse.header.validate", "", int64(34+i), nil)
		}
	}

	switch h.MajorVersion {
	case cfbMajorVersion3:
		if h.SectorShift != cfbSectorShiftV3 {
			return newError(ErrBadHeader, "v3 header must use 512-byte sectors", "parse.header.validate", "", 30, nil)
		}
		if h.NumDirectorySectors != 0 {
			return newError(ErrBadHeader, "v3 header must set directory sectors to zero", "parse.header.validate", "", 40, nil)
		}
	case cfbMajorVersion4:
		if h.SectorShift != cfbSectorShiftV4 {
			return newError(ErrBadHeader, "v4 header must use 4096-byte sectors", "parse.header.validate", "", 30, nil)
		}
	default:
		return newError(ErrBadHeader, fmt.Sprintf("unsupported major version: %d", h.MajorVersion), "parse.header.validate", "", 26, nil)
	}

	return nil
}
