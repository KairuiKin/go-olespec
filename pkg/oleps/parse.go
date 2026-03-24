package oleps

import (
	"encoding/binary"
	"fmt"
	"sort"
	"time"
	"unicode/utf16"
)

func Parse(data []byte) (*Stream, error) {
	if len(data) < 28 {
		return nil, fmt.Errorf("property set stream too small")
	}
	byteOrder := binary.LittleEndian.Uint16(data[0:2])
	if byteOrder != 0xFFFE {
		return nil, fmt.Errorf("invalid byte order: 0x%04X", byteOrder)
	}
	s := &Stream{
		ByteOrder:        byteOrder,
		Version:          binary.LittleEndian.Uint16(data[2:4]),
		SystemIdentifier: binary.LittleEndian.Uint32(data[4:8]),
	}
	copy(s.CLSID[:], data[8:24])
	numSets := binary.LittleEndian.Uint32(data[24:28])
	if numSets == 0 {
		return s, nil
	}
	if len(data) < 28+int(numSets)*20 {
		return nil, fmt.Errorf("property set table truncated")
	}
	s.Sets = make([]PropertySet, 0, numSets)
	for i := 0; i < int(numSets); i++ {
		base := 28 + i*20
		var fmtid GUID
		copy(fmtid[:], data[base:base+16])
		off := binary.LittleEndian.Uint32(data[base+16 : base+20])
		set, err := parseSet(data, off)
		if err != nil {
			return nil, fmt.Errorf("parse set[%d] at %d: %w", i, off, err)
		}
		set.FormatID = fmtid
		s.Sets = append(s.Sets, set)
	}
	return s, nil
}

func parseSet(data []byte, offset uint32) (PropertySet, error) {
	if int(offset) < 0 || int(offset)+8 > len(data) {
		return PropertySet{}, fmt.Errorf("set offset out of bounds: %d", offset)
	}
	section := data[offset:]
	sectionSize := binary.LittleEndian.Uint32(section[0:4])
	propCount := binary.LittleEndian.Uint32(section[4:8])
	if sectionSize < 8 {
		return PropertySet{}, fmt.Errorf("invalid section size: %d", sectionSize)
	}
	if int(offset+sectionSize) > len(data) {
		return PropertySet{}, fmt.Errorf("section size out of bounds")
	}
	section = data[offset : offset+sectionSize]
	if 8+int(propCount)*8 > len(section) {
		return PropertySet{}, fmt.Errorf("property table truncated")
	}

	type entry struct {
		id  uint32
		off uint32
	}
	entries := make([]entry, 0, propCount)
	for i := 0; i < int(propCount); i++ {
		base := 8 + i*8
		id := binary.LittleEndian.Uint32(section[base : base+4])
		off := binary.LittleEndian.Uint32(section[base+4 : base+8])
		if off < 8 || int(off) >= len(section) {
			return PropertySet{}, fmt.Errorf("property offset out of bounds: id=%d off=%d", id, off)
		}
		entries = append(entries, entry{id: id, off: off})
	}
	sort.Slice(entries, func(i, j int) bool {
		a, b := entries[i], entries[j]
		if a.off != b.off {
			return a.off < b.off
		}
		return a.id < b.id
	})

	props := make(map[uint32]Property, len(entries))
	for i, ent := range entries {
		start := int(ent.off)
		end := len(section)
		if i+1 < len(entries) {
			end = int(entries[i+1].off)
		}
		if end < start || end > len(section) {
			return PropertySet{}, fmt.Errorf("invalid property span: id=%d", ent.id)
		}
		raw := section[start:end]
		typ, val, err := parseTypedValue(raw)
		if err != nil {
			return PropertySet{}, fmt.Errorf("property id %d: %w", ent.id, err)
		}
		props[ent.id] = Property{
			ID:    ent.id,
			Type:  typ,
			Value: val,
			Raw:   append([]byte(nil), raw...),
		}
	}
	return PropertySet{
		Properties: props,
		Order:      sortedPropertyIDs(props),
	}, nil
}

func parseTypedValue(raw []byte) (PropertyType, any, error) {
	if len(raw) < 4 {
		return VTEmpty, nil, fmt.Errorf("typed value too small")
	}
	vt := PropertyType(binary.LittleEndian.Uint16(raw[0:2]))
	switch vt {
	case VTEmpty:
		return vt, nil, nil
	case VTI2:
		if len(raw) < 6 {
			return vt, nil, fmt.Errorf("VT_I2 too small")
		}
		return vt, int16(binary.LittleEndian.Uint16(raw[4:6])), nil
	case VTI4:
		if len(raw) < 8 {
			return vt, nil, fmt.Errorf("VT_I4 too small")
		}
		return vt, int32(binary.LittleEndian.Uint32(raw[4:8])), nil
	case VTUI4:
		if len(raw) < 8 {
			return vt, nil, fmt.Errorf("VT_UI4 too small")
		}
		return vt, binary.LittleEndian.Uint32(raw[4:8]), nil
	case VTI8:
		if len(raw) < 12 {
			return vt, nil, fmt.Errorf("VT_I8 too small")
		}
		return vt, int64(binary.LittleEndian.Uint64(raw[4:12])), nil
	case VTUI8:
		if len(raw) < 12 {
			return vt, nil, fmt.Errorf("VT_UI8 too small")
		}
		return vt, binary.LittleEndian.Uint64(raw[4:12]), nil
	case VTBool:
		if len(raw) < 6 {
			return vt, nil, fmt.Errorf("VT_BOOL too small")
		}
		return vt, binary.LittleEndian.Uint16(raw[4:6]) != 0, nil
	case VTLPSTR:
		if len(raw) < 8 {
			return vt, nil, fmt.Errorf("VT_LPSTR too small")
		}
		n := binary.LittleEndian.Uint32(raw[4:8])
		if int(8+n) > len(raw) {
			return vt, nil, fmt.Errorf("VT_LPSTR length out of bounds")
		}
		v := string(raw[8 : 8+n])
		for len(v) > 0 && v[len(v)-1] == 0 {
			v = v[:len(v)-1]
		}
		return vt, v, nil
	case VTLPWSTR:
		if len(raw) < 8 {
			return vt, nil, fmt.Errorf("VT_LPWSTR too small")
		}
		n := binary.LittleEndian.Uint32(raw[4:8]) // number of UTF-16 chars including null
		byteLen := int(n) * 2
		if int(8+byteLen) > len(raw) {
			return vt, nil, fmt.Errorf("VT_LPWSTR length out of bounds")
		}
		u := make([]uint16, n)
		for i := 0; i < int(n); i++ {
			base := 8 + i*2
			u[i] = binary.LittleEndian.Uint16(raw[base : base+2])
		}
		if len(u) > 0 && u[len(u)-1] == 0 {
			u = u[:len(u)-1]
		}
		return vt, string(utf16.Decode(u)), nil
	case VTFiletime:
		if len(raw) < 12 {
			return vt, nil, fmt.Errorf("VT_FILETIME too small")
		}
		ft := binary.LittleEndian.Uint64(raw[4:12])
		return vt, filetimeToUTC(ft), nil
	default:
		return vt, append([]byte(nil), raw[4:]...), nil
	}
}

func filetimeToUTC(ft uint64) time.Time {
	// FILETIME ticks: 100ns since 1601-01-01 UTC
	const unixOffset = int64(11644473600)
	secs := int64(ft/10000000) - unixOffset
	nanos := int64(ft%10000000) * 100
	return time.Unix(secs, nanos).UTC()
}
