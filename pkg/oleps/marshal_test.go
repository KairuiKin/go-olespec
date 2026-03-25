package oleps

import (
	"bytes"
	"testing"
)

func TestMarshalRoundtripSummary(t *testing.T) {
	data := buildSummaryPropertySetStream("Hello OLE", 42)
	s, err := Parse(data)
	if err != nil {
		t.Fatalf("Parse returned error: %v", err)
	}
	set, ok := s.SummaryInformation()
	if !ok {
		t.Fatal("SummaryInformation set not found")
	}
	set.Properties[PIDTitle] = Property{
		ID:    PIDTitle,
		Type:  VTLPWSTR,
		Value: "Edited Title",
	}

	encoded, err := Marshal(s)
	if err != nil {
		t.Fatalf("Marshal returned error: %v", err)
	}
	s2, err := Parse(encoded)
	if err != nil {
		t.Fatalf("Parse(encoded) returned error: %v", err)
	}
	set2, ok := s2.SummaryInformation()
	if !ok {
		t.Fatal("SummaryInformation set not found after roundtrip")
	}
	title, ok := set2.GetString(PIDTitle)
	if !ok || title != "Edited Title" {
		t.Fatalf("unexpected title after roundtrip: %q", title)
	}
	pages, ok := set2.GetInt64(PIDPageCount)
	if !ok || pages != 42 {
		t.Fatalf("unexpected page count after roundtrip: %d", pages)
	}
}

func TestMarshalPreservesUnknownRaw(t *testing.T) {
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
	before, ok := set.Get(pidUnknown)
	if !ok {
		t.Fatal("unknown property not found")
	}

	encoded, err := Marshal(s)
	if err != nil {
		t.Fatalf("Marshal returned error: %v", err)
	}
	s2, err := Parse(encoded)
	if err != nil {
		t.Fatalf("Parse(encoded) returned error: %v", err)
	}
	set2, ok := s2.SummaryInformation()
	if !ok {
		t.Fatal("SummaryInformation set not found after marshal")
	}
	after, ok := set2.Get(pidUnknown)
	if !ok {
		t.Fatal("unknown property missing after marshal")
	}
	if after.Type != vtUnknown {
		t.Fatalf("unexpected unknown property type: 0x%04X", uint16(after.Type))
	}
	if !bytes.Equal(after.Raw, before.Raw) {
		t.Fatalf("unknown property raw changed: before=%X after=%X", before.Raw, after.Raw)
	}
}

func TestMarshalNilStream(t *testing.T) {
	if _, err := Marshal(nil); err == nil {
		t.Fatal("expected error for nil stream")
	}
}

