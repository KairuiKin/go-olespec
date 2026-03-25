package oleps

import (
	"encoding/binary"
	"fmt"
	"sort"
	"time"
)

type GUID [16]byte

func (g GUID) String() string {
	d1 := binary.LittleEndian.Uint32(g[0:4])
	d2 := binary.LittleEndian.Uint16(g[4:6])
	d3 := binary.LittleEndian.Uint16(g[6:8])
	return fmt.Sprintf("%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x",
		d1, d2, d3, g[8], g[9], g[10], g[11], g[12], g[13], g[14], g[15])
}

type PropertyType uint16

const (
	VTEmpty    PropertyType = 0x0000
	VTI2       PropertyType = 0x0002
	VTI4       PropertyType = 0x0003
	VTR4       PropertyType = 0x0004
	VTR8       PropertyType = 0x0005
	VTBool     PropertyType = 0x000B
	VTUI4      PropertyType = 0x0013
	VTI8       PropertyType = 0x0014
	VTUI8      PropertyType = 0x0015
	VTLPSTR    PropertyType = 0x001E
	VTLPWSTR   PropertyType = 0x001F
	VTFiletime PropertyType = 0x0040
)

type Property struct {
	ID    uint32
	Type  PropertyType
	Value any
	Raw   []byte
}

type PropertySet struct {
	FormatID   GUID
	Properties map[uint32]Property
	Order      []uint32
}

func (ps PropertySet) Get(id uint32) (Property, bool) {
	p, ok := ps.Properties[id]
	return p, ok
}

func (ps PropertySet) GetString(id uint32) (string, bool) {
	p, ok := ps.Get(id)
	if !ok {
		return "", false
	}
	v, ok := p.Value.(string)
	return v, ok
}

func (ps PropertySet) GetInt64(id uint32) (int64, bool) {
	p, ok := ps.Get(id)
	if !ok {
		return 0, false
	}
	switch v := p.Value.(type) {
	case int16:
		return int64(v), true
	case int32:
		return int64(v), true
	case int64:
		return v, true
	case uint32:
		return int64(v), true
	case uint64:
		return int64(v), true
	default:
		return 0, false
	}
}

func (ps PropertySet) GetTime(id uint32) (time.Time, bool) {
	p, ok := ps.Get(id)
	if !ok {
		return time.Time{}, false
	}
	v, ok := p.Value.(time.Time)
	return v, ok
}

func (ps *PropertySet) SetString(id uint32, value string) {
	ps.upsert(Property{
		ID:    id,
		Type:  VTLPWSTR,
		Value: value,
	})
}

func (ps *PropertySet) SetInt64(id uint32, value int64) {
	ps.upsert(Property{
		ID:    id,
		Type:  VTI8,
		Value: value,
	})
}

func (ps *PropertySet) SetUint64(id uint32, value uint64) {
	ps.upsert(Property{
		ID:    id,
		Type:  VTUI8,
		Value: value,
	})
}

func (ps *PropertySet) SetBool(id uint32, value bool) {
	ps.upsert(Property{
		ID:    id,
		Type:  VTBool,
		Value: value,
	})
}

func (ps *PropertySet) SetTime(id uint32, value time.Time) {
	ps.upsert(Property{
		ID:    id,
		Type:  VTFiletime,
		Value: value.UTC(),
	})
}

func (ps *PropertySet) Delete(id uint32) {
	if ps == nil || ps.Properties == nil {
		return
	}
	delete(ps.Properties, id)
	ps.Order = sortedPropertyIDs(ps.Properties)
}

func (ps *PropertySet) upsert(p Property) {
	if ps == nil {
		return
	}
	if ps.Properties == nil {
		ps.Properties = map[uint32]Property{}
	}
	p.Raw = nil // recompute on marshal from typed value.
	ps.Properties[p.ID] = p
	ps.Order = sortedPropertyIDs(ps.Properties)
}

type Stream struct {
	ByteOrder        uint16
	Version          uint16
	SystemIdentifier uint32
	CLSID            GUID
	Sets             []PropertySet
}

func (s *Stream) FindSet(formatID GUID) (*PropertySet, bool) {
	if s == nil {
		return nil, false
	}
	for i := range s.Sets {
		if s.Sets[i].FormatID == formatID {
			return &s.Sets[i], true
		}
	}
	return nil, false
}

func (s *Stream) SummaryInformation() (*PropertySet, bool) {
	return s.FindSet(FMTIDSummaryInformation)
}

func (s *Stream) DocumentSummaryInformation() (*PropertySet, bool) {
	return s.FindSet(FMTIDDocumentSummaryInformation)
}

func sortedPropertyIDs(m map[uint32]Property) []uint32 {
	ids := make([]uint32, 0, len(m))
	for id := range m {
		ids = append(ids, id)
	}
	sort.Slice(ids, func(i, j int) bool {
		return ids[i] < ids[j]
	})
	return ids
}
