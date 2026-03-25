package olextract

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
)

func TestExtractBytes(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractBytes(
		buf,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractBytes returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
	if rep.Artifacts[0].Path != "/Docs/A" {
		t.Fatalf("unexpected artifact path: %s", rep.Artifacts[0].Path)
	}
}

func TestExtractFile(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	dir := t.TempDir()
	p := filepath.Join(dir, "sample.cfb")
	if err := os.WriteFile(p, buf, 0o644); err != nil {
		t.Fatalf("WriteFile returned error: %v", err)
	}
	rep, err := ExtractFile(
		p,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractFile returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
}

func TestExtractReader(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractReader(
		bytes.NewReader(buf),
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractReader returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
}

func TestExtractReaderNil(t *testing.T) {
	if _, err := ExtractReader(nil, olecfb.OpenOptions{}, olecfb.ExtractOptions{}); err == nil {
		t.Fatal("expected error for nil reader")
	}
}

func buildSampleCFBBytes(t *testing.T) []byte {
	t.Helper()
	f, err := olecfb.CreateInMemory(olecfb.CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(olecfb.TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.CreateStorage("/Docs"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	if err := tx.PutStream("/Docs/A", bytes.NewReader([]byte("abc")), 3); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if _, err := tx.Commit(nil, olecfb.CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	buf, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	return buf
}
