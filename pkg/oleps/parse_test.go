package oleps

import (
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
