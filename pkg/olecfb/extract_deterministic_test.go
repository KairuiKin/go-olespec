package olecfb

import (
	"bytes"
	"testing"
)

func TestExtractDeterministicRecursiveWithDetectors(t *testing.T) {
	innerBytes, _ := buildValidV4FileWithSingleNormalStream()
	pngPrefix := []byte{0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A}
	oleNative := buildOle10NativeBytes("a.txt", "C:\\a.txt", []byte("abc"))

	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutStream("/Embedded", bytes.NewReader(innerBytes), int64(len(innerBytes))); err != nil {
		t.Fatalf("PutStream /Embedded returned error: %v", err)
	}
	if err := tx.PutStream("/Image1", bytes.NewReader(pngPrefix), int64(len(pngPrefix))); err != nil {
		t.Fatalf("PutStream /Image1 returned error: %v", err)
	}
	if err := tx.PutStream("/OleObj", bytes.NewReader(oleNative), int64(len(oleNative))); err != nil {
		t.Fatalf("PutStream /OleObj returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}

	opt := ExtractOptions{
		IncludeRaw:   true,
		DetectImages: true,
		DetectOLEDS:  true,
		Deduplicate:  false,
		Limits: ExtractLimits{
			MaxDepth:      4,
			MaxArtifacts:  64,
			MaxTotalBytes: 1 << 20,
		},
	}
	r1, err := f.Extract(opt)
	if err != nil {
		t.Fatalf("Extract #1 returned error: %v", err)
	}
	r2, err := f.Extract(opt)
	if err != nil {
		t.Fatalf("Extract #2 returned error: %v", err)
	}

	if len(r1.Artifacts) != len(r2.Artifacts) {
		t.Fatalf("artifact count mismatch: %d vs %d", len(r1.Artifacts), len(r2.Artifacts))
	}
	for i := range r1.Artifacts {
		a, b := r1.Artifacts[i], r2.Artifacts[i]
		if a.Path != b.Path {
			t.Fatalf("path mismatch at %d: %q vs %q", i, a.Path, b.Path)
		}
		if a.Kind != b.Kind {
			t.Fatalf("kind mismatch at %d: %s vs %s", i, a.Kind, b.Kind)
		}
		if a.MediaType != b.MediaType {
			t.Fatalf("media type mismatch at %d: %q vs %q", i, a.MediaType, b.MediaType)
		}
		if a.Note != b.Note {
			t.Fatalf("note mismatch at %d: %q vs %q", i, a.Note, b.Note)
		}
		if a.SHA256 != b.SHA256 {
			t.Fatalf("hash mismatch at %d", i)
		}
		if a.ParentID != b.ParentID {
			t.Fatalf("parent mismatch at %d: %q vs %q", i, a.ParentID, b.ParentID)
		}
		if len(a.Raw) != len(b.Raw) {
			t.Fatalf("raw length mismatch at %d: %d vs %d", i, len(a.Raw), len(b.Raw))
		}
	}
}
