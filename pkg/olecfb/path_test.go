package olecfb

import (
	"strings"
	"testing"
)

func TestParsePathValid(t *testing.T) {
	cases := []string{
		"/",
		"/A",
		"/Folder/Doc",
		"/A~1B",      // decoded segment "A/B"
		"/A~0B",      // decoded segment "A~B"
		"/A~1B/C~0D", // mixed escapes
	}
	for _, tc := range cases {
		p, err := ParsePath(tc)
		if err != nil {
			t.Fatalf("ParsePath(%q) returned error: %v", tc, err)
		}
		if string(p) != tc {
			t.Fatalf("unexpected canonical path for %q: %q", tc, string(p))
		}
	}
}

func TestParsePathInvalid(t *testing.T) {
	tooLongSeg := "/" + strings.Repeat("a", 32)
	tooLongPath := "/" + strings.Repeat("a", 4096)
	cases := []string{
		"",
		"A",
		"/A/",
		"/A//B",
		"/A~",
		"/A~2",
		tooLongSeg,
		tooLongPath,
	}
	for _, tc := range cases {
		if _, err := ParsePath(tc); err == nil {
			t.Fatalf("ParsePath(%q) should fail", tc)
		}
	}
}

func TestJoinPath(t *testing.T) {
	p, err := JoinPath("/", "A/B")
	if err != nil {
		t.Fatalf("JoinPath returned error: %v", err)
	}
	if string(p) != "/A~1B" {
		t.Fatalf("unexpected join result: %q", string(p))
	}
	p2, err := JoinPath("/Root", "A~B")
	if err != nil {
		t.Fatalf("JoinPath returned error: %v", err)
	}
	if string(p2) != "/Root/A~0B" {
		t.Fatalf("unexpected join result: %q", string(p2))
	}
}

func TestParentPathAndBaseName(t *testing.T) {
	if got := string(ParentPath("/")); got != "/" {
		t.Fatalf("unexpected parent for root: %q", got)
	}
	if got := string(ParentPath("/A/B")); got != "/A" {
		t.Fatalf("unexpected parent: %q", got)
	}
	if got := BaseName("/A~1B/C~0D"); got != "C~D" {
		t.Fatalf("unexpected base name: %q", got)
	}
}

func TestDecodeSegmentInvalid(t *testing.T) {
	cases := []string{"~", "~2", "A~"}
	for _, tc := range cases {
		if _, err := DecodeSegment(tc); err == nil {
			t.Fatalf("DecodeSegment(%q) should fail", tc)
		}
	}
}

