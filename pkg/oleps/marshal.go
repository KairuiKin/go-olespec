package oleps

import (
	"encoding/binary"
	"fmt"
	"math"
	"sort"
	"time"
	"unicode/utf16"
)

func Marshal(s *Stream) ([]byte, error) {
	if s == nil {
		return nil, fmt.Errorf("stream is nil")
	}
	setCount := len(s.Sets)
	headerSize := 28 + setCount*20
	sections := make([][]byte, setCount)
	totalSections := 0
	for i := range s.Sets {
		sec, err := marshalSet(s.Sets[i])
		if err != nil {
			return nil, fmt.Errorf("marshal set[%d]: %w", i, err)
		}
		sections[i] = sec
		totalSections += len(sec)
	}

	out := make([]byte, headerSize+totalSections)
	byteOrder := s.ByteOrder
	if byteOrder == 0 {
		byteOrder = 0xFFFE
	}
	binary.LittleEndian.PutUint16(out[0:2], byteOrder)
	binary.LittleEndian.PutUint16(out[2:4], s.Version)
	binary.LittleEndian.PutUint32(out[4:8], s.SystemIdentifier)
	copy(out[8:24], s.CLSID[:])
	binary.LittleEndian.PutUint32(out[24:28], uint32(setCount))

	offset := headerSize
	for i := range s.Sets {
		base := 28 + i*20
		copy(out[base:base+16], s.Sets[i].FormatID[:])
		binary.LittleEndian.PutUint32(out[base+16:base+20], uint32(offset))
		offset += len(sections[i])
	}
	cursor := headerSize
	for i := range sections {
		copy(out[cursor:], sections[i])
		cursor += len(sections[i])
	}
	return out, nil
}

func marshalSet(ps PropertySet) ([]byte, error) {
	ids := marshalPropertyOrder(ps)
	propCount := len(ids)
	headerLen := 8 + propCount*8
	section := make([]byte, headerLen)
	binary.LittleEndian.PutUint32(section[4:8], uint32(propCount))
	offset := headerLen

	for i, id := range ids {
		p, ok := ps.Properties[id]
		if !ok {
			return nil, fmt.Errorf("property %d not found", id)
		}
		tv, err := marshalTypedValue(p)
		if err != nil {
			return nil, fmt.Errorf("property %d: %w", id, err)
		}
		offset = align4(offset)
		section = appendPadding(section, offset)
		base := 8 + i*8
		binary.LittleEndian.PutUint32(section[base:base+4], id)
		binary.LittleEndian.PutUint32(section[base+4:base+8], uint32(offset))
		section = append(section, tv...)
		offset += len(tv)
	}
	section = appendPadding(section, offset)
	binary.LittleEndian.PutUint32(section[0:4], uint32(len(section)))
	return section, nil
}

func marshalPropertyOrder(ps PropertySet) []uint32 {
	if len(ps.Properties) == 0 {
		return nil
	}
	seen := make(map[uint32]struct{}, len(ps.Properties))
	out := make([]uint32, 0, len(ps.Properties))
	for _, id := range ps.Order {
		if _, ok := ps.Properties[id]; !ok {
			continue
		}
		if _, ok := seen[id]; ok {
			continue
		}
		seen[id] = struct{}{}
		out = append(out, id)
	}
	rest := make([]uint32, 0, len(ps.Properties))
	for id := range ps.Properties {
		if _, ok := seen[id]; ok {
			continue
		}
		rest = append(rest, id)
	}
	sort.Slice(rest, func(i, j int) bool { return rest[i] < rest[j] })
	out = append(out, rest...)
	return out
}

func marshalTypedValue(p Property) ([]byte, error) {
	if len(p.Raw) >= 4 {
		return append([]byte(nil), p.Raw...), nil
	}
	buf := make([]byte, 4)
	binary.LittleEndian.PutUint16(buf[0:2], uint16(p.Type))
	switch p.Type {
	case VTEmpty:
		return buf, nil
	case VTI2:
		v, ok := toInt16(p.Value)
		if !ok {
			return nil, fmt.Errorf("VT_I2 expects int16/int32/int64/uint32/uint64")
		}
		out := append(buf, 0, 0)
		binary.LittleEndian.PutUint16(out[4:6], uint16(v))
		return out, nil
	case VTI4:
		v, ok := toInt32(p.Value)
		if !ok {
			return nil, fmt.Errorf("VT_I4 expects int32/int64/int16")
		}
		out := append(buf, make([]byte, 4)...)
		binary.LittleEndian.PutUint32(out[4:8], uint32(v))
		return out, nil
	case VTUI4:
		v, ok := toUint32(p.Value)
		if !ok {
			return nil, fmt.Errorf("VT_UI4 expects uint32/uint64/int64/int32")
		}
		out := append(buf, make([]byte, 4)...)
		binary.LittleEndian.PutUint32(out[4:8], v)
		return out, nil
	case VTI8:
		v, ok := toInt64(p.Value)
		if !ok {
			return nil, fmt.Errorf("VT_I8 expects int64/int32/int16/uint32")
		}
		out := append(buf, make([]byte, 8)...)
		binary.LittleEndian.PutUint64(out[4:12], uint64(v))
		return out, nil
	case VTUI8:
		v, ok := toUint64(p.Value)
		if !ok {
			return nil, fmt.Errorf("VT_UI8 expects uint64/uint32/int64/int32")
		}
		out := append(buf, make([]byte, 8)...)
		binary.LittleEndian.PutUint64(out[4:12], v)
		return out, nil
	case VTBool:
		v, ok := p.Value.(bool)
		if !ok {
			return nil, fmt.Errorf("VT_BOOL expects bool")
		}
		out := append(buf, 0, 0)
		if v {
			binary.LittleEndian.PutUint16(out[4:6], 0xFFFF)
		}
		return out, nil
	case VTLPSTR:
		s, ok := p.Value.(string)
		if !ok {
			return nil, fmt.Errorf("VT_LPSTR expects string")
		}
		out := append(buf, make([]byte, 4)...)
		raw := append([]byte(s), 0)
		binary.LittleEndian.PutUint32(out[4:8], uint32(len(raw)))
		out = append(out, raw...)
		return out, nil
	case VTLPWSTR:
		s, ok := p.Value.(string)
		if !ok {
			return nil, fmt.Errorf("VT_LPWSTR expects string")
		}
		u := utf16.Encode([]rune(s))
		u = append(u, 0)
		out := append(buf, make([]byte, 4)...)
		binary.LittleEndian.PutUint32(out[4:8], uint32(len(u)))
		wide := make([]byte, len(u)*2)
		for i, v := range u {
			binary.LittleEndian.PutUint16(wide[i*2:i*2+2], v)
		}
		out = append(out, wide...)
		return out, nil
	case VTFiletime:
		tm, ok := p.Value.(time.Time)
		if !ok {
			return nil, fmt.Errorf("VT_FILETIME expects time.Time")
		}
		out := append(buf, make([]byte, 8)...)
		binary.LittleEndian.PutUint64(out[4:12], utcToFiletime(tm))
		return out, nil
	default:
		if b, ok := p.Value.([]byte); ok {
			return append(buf, b...), nil
		}
		return nil, fmt.Errorf("unsupported type 0x%04X without raw bytes", uint16(p.Type))
	}
}

func toInt16(v any) (int16, bool) {
	switch x := v.(type) {
	case int16:
		return x, true
	case int32:
		if x < math.MinInt16 || x > math.MaxInt16 {
			return 0, false
		}
		return int16(x), true
	case int64:
		if x < math.MinInt16 || x > math.MaxInt16 {
			return 0, false
		}
		return int16(x), true
	case uint32:
		if x > math.MaxInt16 {
			return 0, false
		}
		return int16(x), true
	case uint64:
		if x > math.MaxInt16 {
			return 0, false
		}
		return int16(x), true
	default:
		return 0, false
	}
}

func toInt32(v any) (int32, bool) {
	switch x := v.(type) {
	case int16:
		return int32(x), true
	case int32:
		return x, true
	case int64:
		if x < math.MinInt32 || x > math.MaxInt32 {
			return 0, false
		}
		return int32(x), true
	case uint32:
		if x > math.MaxInt32 {
			return 0, false
		}
		return int32(x), true
	case uint64:
		if x > math.MaxInt32 {
			return 0, false
		}
		return int32(x), true
	default:
		return 0, false
	}
}

func toInt64(v any) (int64, bool) {
	switch x := v.(type) {
	case int16:
		return int64(x), true
	case int32:
		return int64(x), true
	case int64:
		return x, true
	case uint32:
		return int64(x), true
	case uint64:
		if x > math.MaxInt64 {
			return 0, false
		}
		return int64(x), true
	default:
		return 0, false
	}
}

func toUint32(v any) (uint32, bool) {
	switch x := v.(type) {
	case int16:
		if x < 0 {
			return 0, false
		}
		return uint32(x), true
	case int32:
		if x < 0 {
			return 0, false
		}
		return uint32(x), true
	case int64:
		if x < 0 || x > math.MaxUint32 {
			return 0, false
		}
		return uint32(x), true
	case uint32:
		return x, true
	case uint64:
		if x > math.MaxUint32 {
			return 0, false
		}
		return uint32(x), true
	default:
		return 0, false
	}
}

func toUint64(v any) (uint64, bool) {
	switch x := v.(type) {
	case int16:
		if x < 0 {
			return 0, false
		}
		return uint64(x), true
	case int32:
		if x < 0 {
			return 0, false
		}
		return uint64(x), true
	case int64:
		if x < 0 {
			return 0, false
		}
		return uint64(x), true
	case uint32:
		return uint64(x), true
	case uint64:
		return x, true
	default:
		return 0, false
	}
}

func utcToFiletime(t time.Time) uint64 {
	const unixOffset = int64(11644473600)
	u := t.UTC()
	secs := u.Unix() + unixOffset
	return uint64(secs*10000000 + int64(u.Nanosecond()/100))
}

func align4(n int) int {
	return (n + 3) &^ 3
}

func appendPadding(buf []byte, wantLen int) []byte {
	if len(buf) >= wantLen {
		return buf
	}
	return append(buf, make([]byte, wantLen-len(buf))...)
}
