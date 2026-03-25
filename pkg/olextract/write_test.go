package olextract

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
)

func TestWriteArtifacts(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractBytes(
		buf,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true, IncludeRaw: true},
	)
	if err != nil {
		t.Fatalf("ExtractBytes returned error: %v", err)
	}
	outDir := t.TempDir()
	res, err := WriteArtifacts(rep, outDir, WriteOptions{})
	if err != nil {
		t.Fatalf("WriteArtifacts returned error: %v", err)
	}
	if res.FilesWritten != 1 {
		t.Fatalf("unexpected files written: %d", res.FilesWritten)
	}
	if len(res.Files) != 1 {
		t.Fatalf("unexpected files list count: %d", len(res.Files))
	}
	got, err := os.ReadFile(res.Files[0].FilePath)
	if err != nil {
		t.Fatalf("ReadFile returned error: %v", err)
	}
	if string(got) != "abc" {
		t.Fatalf("unexpected extracted payload: %q", string(got))
	}
}

func TestWriteArtifactsSkipNoRaw(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractBytes(
		buf,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true, IncludeRaw: false},
	)
	if err != nil {
		t.Fatalf("ExtractBytes returned error: %v", err)
	}
	res, err := WriteArtifacts(rep, t.TempDir(), WriteOptions{})
	if err != nil {
		t.Fatalf("WriteArtifacts returned error: %v", err)
	}
	if res.FilesWritten != 0 {
		t.Fatalf("unexpected files written: %d", res.FilesWritten)
	}
	if res.Skipped != 1 {
		t.Fatalf("unexpected skipped count: %d", res.Skipped)
	}
}

func TestWriteArtifactsConflict(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractBytes(
		buf,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true, IncludeRaw: true},
	)
	if err != nil {
		t.Fatalf("ExtractBytes returned error: %v", err)
	}
	outDir := t.TempDir()
	if _, err := WriteArtifacts(rep, outDir, WriteOptions{}); err != nil {
		t.Fatalf("first WriteArtifacts returned error: %v", err)
	}
	if _, err := WriteArtifacts(rep, outDir, WriteOptions{}); err == nil {
		t.Fatal("expected conflict on second write")
	} else if !olecfb.IsCode(err, olecfb.ErrConflict) {
		t.Fatalf("expected ErrConflict, got %v", err)
	}
}

func TestExtractFileToDir(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	src := filepath.Join(t.TempDir(), "sample.cfb")
	if err := os.WriteFile(src, buf, 0o644); err != nil {
		t.Fatalf("WriteFile returned error: %v", err)
	}
	outDir := t.TempDir()
	rep, writeRes, err := ExtractFileToDir(
		src,
		outDir,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true, IncludeRaw: false},
		WriteOptions{},
	)
	if err != nil {
		t.Fatalf("ExtractFileToDir returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
	if writeRes.FilesWritten != 1 {
		t.Fatalf("unexpected files written: %d", writeRes.FilesWritten)
	}
	if _, err := os.Stat(writeRes.Files[0].FilePath); err != nil {
		t.Fatalf("expected output file: %v", err)
	}
}

func TestExtractBytesToDir(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, writeRes, err := ExtractBytesToDir(
		buf,
		t.TempDir(),
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true, IncludeRaw: false},
		WriteOptions{},
	)
	if err != nil {
		t.Fatalf("ExtractBytesToDir returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 || writeRes.FilesWritten != 1 {
		t.Fatalf("unexpected result: artifacts=%d files=%d", rep.Stats.ArtifactsTotal, writeRes.FilesWritten)
	}
}

func TestExtractReaderToDir(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, writeRes, err := ExtractReaderToDir(
		bytes.NewReader(buf),
		t.TempDir(),
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
		WriteOptions{},
	)
	if err != nil {
		t.Fatalf("ExtractReaderToDir returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 || writeRes.FilesWritten != 1 {
		t.Fatalf("unexpected result: artifacts=%d files=%d", rep.Stats.ArtifactsTotal, writeRes.FilesWritten)
	}
}

func TestExtractBackendToDir(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, writeRes, err := ExtractBackendToDir(
		&testReadBackend{buf: buf},
		t.TempDir(),
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
		WriteOptions{},
	)
	if err != nil {
		t.Fatalf("ExtractBackendToDir returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 || writeRes.FilesWritten != 1 {
		t.Fatalf("unexpected result: artifacts=%d files=%d", rep.Stats.ArtifactsTotal, writeRes.FilesWritten)
	}
}
