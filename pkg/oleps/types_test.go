package oleps

import (
	"testing"
	"time"
)

func TestPropertySetSettersAndDelete(t *testing.T) {
	ps := &PropertySet{FormatID: FMTIDSummaryInformation}
	ts := time.Date(2026, 3, 1, 2, 3, 4, 0, time.UTC)

	ps.SetString(PIDTitle, "hello")
	ps.SetInt64(PIDPageCount, 7)
	ps.SetBool(PIDSecurity, true)
	ps.SetTime(PIDLastSaveTime, ts)
	ps.SetUint64(0x9000, 99)

	if len(ps.Properties) != 5 {
		t.Fatalf("unexpected property count: %d", len(ps.Properties))
	}
	if got, ok := ps.GetString(PIDTitle); !ok || got != "hello" {
		t.Fatalf("unexpected title: %q", got)
	}
	if got, ok := ps.GetInt64(PIDPageCount); !ok || got != 7 {
		t.Fatalf("unexpected page count: %d", got)
	}
	if p, ok := ps.Get(PIDSecurity); !ok || p.Type != VTBool {
		t.Fatalf("unexpected security property")
	}
	if got, ok := ps.GetTime(PIDLastSaveTime); !ok || !got.Equal(ts) {
		t.Fatalf("unexpected time value: %v", got)
	}

	ps.Delete(PIDPageCount)
	if _, ok := ps.Get(PIDPageCount); ok {
		t.Fatal("deleted property still exists")
	}
}

func TestPropertySetSettersRoundtrip(t *testing.T) {
	s := &Stream{
		ByteOrder: 0xFFFE,
		Sets: []PropertySet{
			{FormatID: FMTIDSummaryInformation},
		},
	}
	ps := &s.Sets[0]
	ps.SetString(PIDTitle, "roundtrip")
	ps.SetInt64(PIDPageCount, 11)
	ps.SetBool(PIDSecurity, false)

	enc, err := Marshal(s)
	if err != nil {
		t.Fatalf("Marshal returned error: %v", err)
	}
	s2, err := Parse(enc)
	if err != nil {
		t.Fatalf("Parse returned error: %v", err)
	}
	ps2, ok := s2.SummaryInformation()
	if !ok {
		t.Fatal("summary set not found")
	}
	if title, ok := ps2.GetString(PIDTitle); !ok || title != "roundtrip" {
		t.Fatalf("unexpected title: %q", title)
	}
	if pc, ok := ps2.GetInt64(PIDPageCount); !ok || pc != 11 {
		t.Fatalf("unexpected page count: %d", pc)
	}
}

