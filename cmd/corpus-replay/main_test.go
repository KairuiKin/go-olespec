package main

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
)

func TestRunReplayMixedCorpus(t *testing.T) {
	root := t.TempDir()
	validPath := filepath.Join(root, "ok.cfb")
	if err := os.WriteFile(validPath, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile valid returned error: %v", err)
	}
	invalidPath := filepath.Join(root, "bad.cfb")
	if err := os.WriteFile(invalidPath, []byte("not-cfb"), 0o644); err != nil {
		t.Fatalf("WriteFile invalid returned error: %v", err)
	}

	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-mode", "lenient"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}

	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 2 {
		t.Fatalf("unexpected processed: %d", rep.Summary.Processed)
	}
	if rep.Summary.Success != 1 {
		t.Fatalf("unexpected success: %d", rep.Summary.Success)
	}
	if rep.Summary.Failed != 1 {
		t.Fatalf("unexpected failed: %d", rep.Summary.Failed)
	}
	if len(rep.Files) != 2 {
		t.Fatalf("unexpected file entries: %d", len(rep.Files))
	}
}

func TestRunReplayOutputFile(t *testing.T) {
	root := t.TempDir()
	validPath := filepath.Join(root, "ok.cfb")
	if err := os.WriteFile(validPath, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile valid returned error: %v", err)
	}
	output := filepath.Join(root, "out", "report.json")
	if err := run([]string{"-root", root, "-ext", ".cfb", "-output", output}, &bytes.Buffer{}); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	buf, err := os.ReadFile(output)
	if err != nil {
		t.Fatalf("ReadFile returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(buf, &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 1 || rep.Summary.Success != 1 {
		t.Fatalf("unexpected summary: %+v", rep.Summary)
	}
}

func TestParseModeInvalid(t *testing.T) {
	if _, err := parseMode("bad"); err == nil {
		t.Fatal("expected parseMode error")
	}
}

func buildSampleCFB(t *testing.T) []byte {
	t.Helper()
	f, err := olecfb.CreateInMemory(olecfb.CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	defer f.Close()
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
