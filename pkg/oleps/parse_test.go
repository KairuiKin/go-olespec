package oleps

import (
	"bytes"
	"encoding/binary"
	"testing"
)

func TestParseSummaryInformation(t *testing.T) {
	data := buildSummaryPropertySetStream("Hello OLE", 42)
	s, err := Parse(data)
	if err != nil {
		t.Fatalf("Parse returned error: %v", err)
	}
	set, ok := s.SummaryInformation()
	if !ok {
		t.Fatal("SummaryInformation set not found")
	}
	title, ok := set.GetString(PIDTitle)
	if !ok {
		t.Fatal("title property not found")
	}
	if title != "Hello OLE" {
		t.Fatalf("unexpected title: %q", title)
	}
	pageCount, ok := set.GetInt64(PIDPageCount)
	if !ok {
		t.Fatal("page count property not found")
	}
	if pageCount != 42 {
		t.Fatalf("unexpected page count: %d", pageCount)
	}
}

func TestParseInvalidByteOrder(t *testing.T) {
	data := buildSummaryPropertySetStream("X", 1)
	data[0] = 0
	if _, err := Parse(data); err == nil {
		t.Fatal("expected parse failure for invalid byte order")
	}
}

func TestParseUnknownPropertyPreserved(t *testing.T) {
	const (
		pidUnknown = uint32(0x777)
		vtUnknown  = PropertyType(0x1337)
	)
	data := buildPropertySetWithUnknown(pidUnknown, vtUnknown, []byte{0xDE, 0xAD, 0xBE, 0xEF})
	s, err := Parse(data)
	if err != nil {
		t.Fatalf("Parse returned error: %v", err)
	}
	set, ok := s.SummaryInformation()
	if !ok {
		t.Fatal("SummaryInformation set not found")
	}
	p, ok := set.Get(pidUnknown)
	if !ok {
		t.Fatal("unknown property not found")
	}
	if p.Type != vtUnknown {
		t.Fatalf("unexpected property type: 0x%04X", uint16(p.Type))
	}
	v, ok := p.Value.([]byte)
	if !ok {
		t.Fatalf("unexpected property value type: %T", p.Value)
	}
	if !bytes.Equal(v, []byte{0xDE, 0xAD, 0xBE, 0xEF}) {
		t.Fatalf("unexpected property value: %X", v)
	}
	wantRaw := []byte{
		0x37, 0x13, 0x00, 0x00, // VT + reserved
		0xDE, 0xAD, 0xBE, 0xEF, // payload
	}
	if !bytes.Equal(p.Raw, wantRaw) {
		t.Fatalf("unexpected raw typed value: %X", p.Raw)
	}
}

func buildSummaryPropertySetStream(title string, pageCount int32) []byte {
	// Property set stream header (28 bytes) + one property set descriptor (20 bytes).
	headerSize := 28 + 20
	header := make([]byte, headerSize)
	binary.LittleEndian.PutUint16(header[0:2], 0xFFFE) // byte order
	binary.LittleEndian.PutUint16(header[2:4], 0x0000) // version
	binary.LittleEndian.PutUint32(header[4:8], 0x00000000)
	// CLSID [8:24] remains zero
	binary.LittleEndian.PutUint32(header[24:28], 1) // number of property sets
	copy(header[28:44], FMTIDSummaryInformation[:])
	binary.LittleEndian.PutUint32(header[44:48], uint32(headerSize))

	titleUTF16 := utf16LE(title + "\x00")
	valTitle := make([]byte, 8+len(titleUTF16))
	binary.LittleEndian.PutUint16(valTitle[0:2], uint16(VTLPWSTR))
	binary.LittleEndian.PutUint16(valTitle[2:4], 0)
	binary.LittleEndian.PutUint32(valTitle[4:8], uint32(len(title)+1))
	copy(valTitle[8:], titleUTF16)

	valPages := make([]byte, 8)
	binary.LittleEndian.PutUint16(valPages[0:2], uint16(VTI4))
	binary.LittleEndian.PutUint16(valPages[2:4], 0)
	binary.LittleEndian.PutUint32(valPages[4:8], uint32(pageCount))

	// Section layout:
	// [0:4] size
	// [4:8] property count
	// [8:16]  entry 1
	// [16:24] entry 2
	// [24:..] values
	offTitle := uint32(24)
	offPages := offTitle + uint32(len(valTitle))
	sectionSize := int(offPages) + len(valPages)

	section := make([]byte, sectionSize)
	binary.LittleEndian.PutUint32(section[0:4], uint32(sectionSize))
	binary.LittleEndian.PutUint32(section[4:8], 2)
	binary.LittleEndian.PutUint32(section[8:12], PIDTitle)
	binary.LittleEndian.PutUint32(section[12:16], offTitle)
	binary.LittleEndian.PutUint32(section[16:20], PIDPageCount)
	binary.LittleEndian.PutUint32(section[20:24], offPages)
	copy(section[offTitle:], valTitle)
	copy(section[offPages:], valPages)

	return append(header, section...)
}

func buildPropertySetWithUnknown(pid uint32, vt PropertyType, payload []byte) []byte {
	headerSize := 28 + 20
	header := make([]byte, headerSize)
	binary.LittleEndian.PutUint16(header[0:2], 0xFFFE)
	binary.LittleEndian.PutUint16(header[2:4], 0x0000)
	binary.LittleEndian.PutUint32(header[24:28], 1)
	copy(header[28:44], FMTIDSummaryInformation[:])
	binary.LittleEndian.PutUint32(header[44:48], uint32(headerSize))

	val := make([]byte, 4+len(payload))
	binary.LittleEndian.PutUint16(val[0:2], uint16(vt))
	copy(val[4:], payload)

	off := uint32(16)
	sectionSize := int(off) + len(val)
	section := make([]byte, sectionSize)
	binary.LittleEndian.PutUint32(section[0:4], uint32(sectionSize))
	binary.LittleEndian.PutUint32(section[4:8], 1)
	binary.LittleEndian.PutUint32(section[8:12], pid)
	binary.LittleEndian.PutUint32(section[12:16], off)
	copy(section[off:], val)

	return append(header, section...)
}

func utf16LE(s string) []byte {
	r := []rune(s)
	u := make([]uint16, len(r))
	for i := range r {
		u[i] = uint16(r[i])
	}
	out := make([]byte, len(u)*2)
	for i, v := range u {
		binary.LittleEndian.PutUint16(out[i*2:i*2+2], v)
	}
	return out
}
