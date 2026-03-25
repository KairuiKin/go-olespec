package olextract

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
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
	if res.Files[0].RelativePath == "" {
		t.Fatal("expected relative path")
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

func TestWriteArtifactsConflictNoPartialWrites(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{Path: "/A", Kind: olecfb.ArtifactStream, Raw: []byte("aaa")},
			{Path: "/B", Kind: olecfb.ArtifactStream, Raw: []byte("bbb")},
		},
	}
	outDir := t.TempDir()
	// Pre-create the second target to trigger conflict.
	if err := os.WriteFile(filepath.Join(outDir, "000001_B.bin"), []byte("x"), 0o644); err != nil {
		t.Fatalf("WriteFile returned error: %v", err)
	}
	if _, err := WriteArtifacts(rep, outDir, WriteOptions{}); err == nil {
		t.Fatal("expected conflict")
	} else if !olecfb.IsCode(err, olecfb.ErrConflict) {
		t.Fatalf("expected ErrConflict, got %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "000000_A.bin")); err == nil {
		t.Fatal("unexpected partial write for first file")
	} else if !os.IsNotExist(err) {
		t.Fatalf("unexpected stat error: %v", err)
	}
}

func TestWriteArtifactsTreeLayoutAndManifest(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{
				ID:     "a1",
				Path:   "/Docs/A",
				Kind:   olecfb.ArtifactStream,
				Raw:    []byte("aaa"),
				Size:   3,
				SHA256: "h1",
			},
			{
				ID:     "a2",
				Path:   "/Embedded!/Blob",
				Kind:   olecfb.ArtifactOLEFile,
				Raw:    []byte("bbb"),
				Size:   3,
				SHA256: "h2",
			},
		},
	}
	outDir := t.TempDir()
	res, err := WriteArtifacts(rep, outDir, WriteOptions{
		Layout:        WriteLayoutTree,
		WriteManifest: true,
	})
	if err != nil {
		t.Fatalf("WriteArtifacts returned error: %v", err)
	}
	if res.FilesWritten != 2 {
		t.Fatalf("unexpected files written: %d", res.FilesWritten)
	}
	if res.ManifestPath == "" {
		t.Fatal("expected manifest path")
	}
	for _, f := range res.Files {
		if _, err := os.Stat(f.FilePath); err != nil {
			t.Fatalf("written file missing: %s", f.FilePath)
		}
	}
	mf, err := os.ReadFile(res.ManifestPath)
	if err != nil {
		t.Fatalf("ReadFile manifest returned error: %v", err)
	}
	var parsed struct {
		Files []struct {
			ArtifactID   string `json:"artifact_id"`
			RelativePath string `json:"relative_path"`
		} `json:"files"`
	}
	if err := json.Unmarshal(mf, &parsed); err != nil {
		t.Fatalf("Unmarshal manifest returned error: %v", err)
	}
	if len(parsed.Files) != 2 {
		t.Fatalf("unexpected manifest files count: %d", len(parsed.Files))
	}
	for _, f := range parsed.Files {
		if f.RelativePath == "" {
			t.Fatal("expected manifest relative_path")
		}
	}
}

func TestWriteArtifactsInvalidLayout(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{Path: "/A", Raw: []byte("a"), Kind: olecfb.ArtifactStream},
		},
	}
	if _, err := WriteArtifacts(rep, t.TempDir(), WriteOptions{Layout: WriteLayout("bad")}); err == nil {
		t.Fatal("expected invalid layout error")
	} else if !olecfb.IsCode(err, olecfb.ErrInvalidArgument) {
		t.Fatalf("expected ErrInvalidArgument, got %v", err)
	}
}

func TestWriteArtifactsManifestConflict(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{Path: "/A", Kind: olecfb.ArtifactStream}, // no raw: only manifest output.
		},
	}
	outDir := t.TempDir()
	opt := WriteOptions{WriteManifest: true}
	if _, err := WriteArtifacts(rep, outDir, opt); err != nil {
		t.Fatalf("first WriteArtifacts returned error: %v", err)
	}
	if _, err := WriteArtifacts(rep, outDir, opt); err == nil {
		t.Fatal("expected manifest conflict")
	} else if !olecfb.IsCode(err, olecfb.ErrConflict) {
		t.Fatalf("expected ErrConflict, got %v", err)
	}
}

func TestWriteArtifactsManifestConflictNoPartialWrites(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{Path: "/A", Kind: olecfb.ArtifactStream, Raw: []byte("aaa")},
		},
	}
	outDir := t.TempDir()
	if err := os.WriteFile(filepath.Join(outDir, "manifest.json"), []byte("{}"), 0o644); err != nil {
		t.Fatalf("WriteFile returned error: %v", err)
	}
	if _, err := WriteArtifacts(rep, outDir, WriteOptions{WriteManifest: true}); err == nil {
		t.Fatal("expected manifest conflict")
	} else if !olecfb.IsCode(err, olecfb.ErrConflict) {
		t.Fatalf("expected ErrConflict, got %v", err)
	}
	if _, err := os.Stat(filepath.Join(outDir, "000000_A.bin")); err == nil {
		t.Fatal("unexpected partial write before manifest conflict")
	} else if !os.IsNotExist(err) {
		t.Fatalf("unexpected stat error: %v", err)
	}
}

func TestWriteArtifactsManifestNameTraversalRejected(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{Path: "/A", Kind: olecfb.ArtifactStream, Raw: []byte("a")},
		},
	}
	_, err := WriteArtifacts(rep, t.TempDir(), WriteOptions{
		WriteManifest: true,
		ManifestName:  "../manifest.json",
	})
	if err == nil {
		t.Fatal("expected invalid manifest name error")
	}
	if !olecfb.IsCode(err, olecfb.ErrInvalidArgument) {
		t.Fatalf("expected ErrInvalidArgument, got %v", err)
	}
}

func TestWriteArtifactsUsesOleFileNameExtension(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{
				ID:          "a1",
				Path:        "/Obj",
				Kind:        olecfb.ArtifactOleObj,
				OLEFileName: "hello.TXT",
				Raw:         []byte("abc"),
			},
		},
	}
	res, err := WriteArtifacts(rep, t.TempDir(), WriteOptions{})
	if err != nil {
		t.Fatalf("WriteArtifacts returned error: %v", err)
	}
	if res.FilesWritten != 1 {
		t.Fatalf("unexpected files written: %d", res.FilesWritten)
	}
	if !strings.HasSuffix(strings.ToLower(res.Files[0].FilePath), ".txt") {
		t.Fatalf("expected .txt suffix, got %s", res.Files[0].FilePath)
	}
}

func TestWriteArtifactsTreeLayoutAvoidsReservedWindowsNames(t *testing.T) {
	rep := &olecfb.ExtractReport{
		Artifacts: []olecfb.Artifact{
			{
				Path: "/CON/A",
				Kind: olecfb.ArtifactStream,
				Raw:  []byte("abc"),
			},
		},
	}
	res, err := WriteArtifacts(rep, t.TempDir(), WriteOptions{Layout: WriteLayoutTree})
	if err != nil {
		t.Fatalf("WriteArtifacts returned error: %v", err)
	}
	if len(res.Files) != 1 {
		t.Fatalf("unexpected files count: %d", len(res.Files))
	}
	rel := strings.ToUpper(filepath.ToSlash(res.Files[0].RelativePath))
	if strings.HasPrefix(rel, "CON/") {
		t.Fatalf("reserved path segment should be rewritten, got %s", rel)
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
