package main

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
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

func TestRunReplayBaselineDiffAndNewlyFailedGate(t *testing.T) {
	root := t.TempDir()
	target := filepath.Join(root, "target.cfb")
	if err := os.WriteFile(target, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile target returned error: %v", err)
	}
	baselinePath := filepath.Join(root, "baseline.json")
	if err := run([]string{"-root", root, "-ext", ".cfb", "-output", baselinePath}, &bytes.Buffer{}); err != nil {
		t.Fatalf("baseline run returned error: %v", err)
	}
	if err := os.WriteFile(target, []byte("bad"), 0o644); err != nil {
		t.Fatalf("overwrite target returned error: %v", err)
	}

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-baseline", baselinePath,
		"-max-newly-failed", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected gate failure")
	}
	if !strings.Contains(err.Error(), "newly_failed") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Baseline == nil {
		t.Fatal("expected baseline diff in report")
	}
	if rep.Baseline.NewlyFailed != 1 {
		t.Fatalf("unexpected newly failed: %d", rep.Baseline.NewlyFailed)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate to fail")
	}
}

func TestRunReplayMaxFailedGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-max-failed", "0"}, &out)
	if err == nil {
		t.Fatal("expected max-failed gate error")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.Failed != 1 {
		t.Fatalf("unexpected failed count: %d", rep.Summary.Failed)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxNewlyFailedRequiresBaseline(t *testing.T) {
	var out bytes.Buffer
	err := run([]string{"-max-newly-failed", "0"}, &out)
	if err == nil {
		t.Fatal("expected validation error")
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
