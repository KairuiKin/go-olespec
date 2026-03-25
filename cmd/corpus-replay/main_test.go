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
	if len(rep.Summary.ErrorCodes) == 0 {
		t.Fatal("expected error code summary")
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

func TestRunReplayDenyErrorCodesGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}

	var probe bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb"}, &probe); err != nil {
		t.Fatalf("probe run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(probe.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal probe returned error: %v", err)
	}
	code := ""
	for k := range rep.Summary.ErrorCodes {
		code = k
		break
	}
	if code == "" {
		t.Fatal("expected at least one error code")
	}

	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-deny-error-codes", code}, &out)
	if err == nil {
		t.Fatal("expected deny-error-codes gate failure")
	}
	var rep2 replayReport
	if err := json.Unmarshal(out.Bytes(), &rep2); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep2.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayErrorCodeBaselineGates(t *testing.T) {
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
		"-max-new-error-codes", "0",
		"-max-error-code-regressions", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected error-code baseline gate failure")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Baseline == nil {
		t.Fatal("expected baseline")
	}
	if len(rep.Baseline.NewErrorCodes) == 0 {
		t.Fatal("expected new error codes")
	}
	if len(rep.Baseline.IncreasedErrorCodes) == 0 {
		t.Fatal("expected increased error codes")
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayErrorCodeGatesRequireBaseline(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-max-new-error-codes", "0"}, &out); err == nil {
		t.Fatal("expected max-new-error-codes validation error")
	}
	if err := run([]string{"-max-error-code-regressions", "0"}, &out); err == nil {
		t.Fatal("expected max-error-code-regressions validation error")
	}
}

func TestRunReplayBaselineFileSetGates(t *testing.T) {
	root := t.TempDir()
	a := filepath.Join(root, "a.cfb")
	if err := os.WriteFile(a, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	baselinePath := filepath.Join(root, "baseline.json")
	if err := run([]string{"-root", root, "-ext", ".cfb", "-output", baselinePath}, &bytes.Buffer{}); err != nil {
		t.Fatalf("baseline run returned error: %v", err)
	}
	b := filepath.Join(root, "b.cfb")
	if err := os.WriteFile(b, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile b returned error: %v", err)
	}

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-baseline", baselinePath,
		"-max-new-files", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-new-files gate failure")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Baseline == nil || rep.Baseline.NewFiles == 0 {
		t.Fatalf("unexpected baseline new files: %+v", rep.Baseline)
	}
}

func TestRunReplayBaselineRemovedFilesGate(t *testing.T) {
	root := t.TempDir()
	a := filepath.Join(root, "a.cfb")
	b := filepath.Join(root, "b.cfb")
	if err := os.WriteFile(a, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	if err := os.WriteFile(b, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile b returned error: %v", err)
	}
	baselinePath := filepath.Join(root, "baseline.json")
	if err := run([]string{"-root", root, "-ext", ".cfb", "-output", baselinePath}, &bytes.Buffer{}); err != nil {
		t.Fatalf("baseline run returned error: %v", err)
	}
	if err := os.Remove(b); err != nil {
		t.Fatalf("Remove b returned error: %v", err)
	}

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-baseline", baselinePath,
		"-max-removed-files", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-removed-files gate failure")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Baseline == nil || rep.Baseline.RemovedFiles == 0 {
		t.Fatalf("unexpected baseline removed files: %+v", rep.Baseline)
	}
}

func TestRunReplayBaselineNewlyPartialGate(t *testing.T) {
	root := t.TempDir()
	a := filepath.Join(root, "a.cfb")
	if err := os.WriteFile(a, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	baselinePath := filepath.Join(root, "baseline.json")
	if err := run([]string{"-root", root, "-ext", ".cfb", "-output", baselinePath}, &bytes.Buffer{}); err != nil {
		t.Fatalf("baseline run returned error: %v", err)
	}

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-baseline", baselinePath,
		"-max-artifact-size", "1",
		"-max-newly-partial", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-newly-partial gate failure")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Baseline == nil || rep.Baseline.NewlyPartial == 0 {
		t.Fatalf("unexpected baseline newly partial: %+v", rep.Baseline)
	}
}

func TestRunReplayBaselineFileSetGatesRequireBaseline(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-max-new-files", "0"}, &out); err == nil {
		t.Fatal("expected max-new-files validation error")
	}
	if err := run([]string{"-max-removed-files", "0"}, &out); err == nil {
		t.Fatal("expected max-removed-files validation error")
	}
	if err := run([]string{"-max-newly-partial", "0"}, &out); err == nil {
		t.Fatal("expected max-newly-partial validation error")
	}
}

func TestRunReplayTrendSummary(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	if err := os.MkdirAll(history, 0o755); err != nil {
		t.Fatalf("MkdirAll history returned error: %v", err)
	}
	writeTrendReport(t, filepath.Join(history, "r1.json"), replayReport{
		GeneratedAt: "2026-01-01T00:00:00Z",
		Options:     replayOptions{RunID: "h1"},
		Summary: replaySummary{
			Processed: 10,
			Success:   10,
			Failed:    0,
			Partial:   0,
			PassRate:  1.0,
		},
	})
	writeTrendReport(t, filepath.Join(history, "r2.json"), replayReport{
		GeneratedAt: "2026-01-02T00:00:00Z",
		Options:     replayOptions{RunID: "h2"},
		Summary: replaySummary{
			Processed: 10,
			Success:   9,
			Failed:    1,
			Partial:   0,
			PassRate:  0.9,
		},
	})

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-run-id", "cur",
		"-trend-dir", history,
	}, &out); err != nil {
		t.Fatalf("run with trend returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Trend == nil {
		t.Fatal("expected trend summary")
	}
	if len(rep.Trend.Points) != 3 {
		t.Fatalf("unexpected trend points: %d", len(rep.Trend.Points))
	}
	last := rep.Trend.Points[len(rep.Trend.Points)-1]
	if last.RunID != "cur" || last.ReportPath != "current" {
		t.Fatalf("unexpected current point: %+v", last)
	}
	if rep.Trend.LatestDelta == nil {
		t.Fatal("expected latest delta")
	}
}

func TestRunReplayTrendLimit(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	if err := os.MkdirAll(history, 0o755); err != nil {
		t.Fatalf("MkdirAll history returned error: %v", err)
	}
	writeTrendReport(t, filepath.Join(history, "r1.json"), replayReport{
		GeneratedAt: "2026-01-01T00:00:00Z",
		Options:     replayOptions{RunID: "h1"},
		Summary:     replaySummary{Processed: 1, Success: 1, PassRate: 1},
	})
	writeTrendReport(t, filepath.Join(history, "r2.json"), replayReport{
		GeneratedAt: "2026-01-02T00:00:00Z",
		Options:     replayOptions{RunID: "h2"},
		Summary:     replaySummary{Processed: 1, Success: 1, PassRate: 1},
	})

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-run-id", "cur",
		"-trend-dir", history,
		"-trend-limit", "1",
	}, &out); err != nil {
		t.Fatalf("run with trend limit returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Trend == nil {
		t.Fatal("expected trend summary")
	}
	if len(rep.Trend.Points) != 2 {
		t.Fatalf("unexpected trend points count: %d", len(rep.Trend.Points))
	}
	if rep.Trend.Points[0].RunID != "h2" {
		t.Fatalf("expected latest history point to be h2, got %+v", rep.Trend.Points[0])
	}
}

func TestRunReplayTrendGates(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	if err := os.MkdirAll(history, 0o755); err != nil {
		t.Fatalf("MkdirAll history returned error: %v", err)
	}
	writeTrendReport(t, filepath.Join(history, "r1.json"), replayReport{
		GeneratedAt: "2026-01-01T00:00:00Z",
		Options:     replayOptions{RunID: "h1"},
		Summary:     replaySummary{Processed: 1, Success: 1, Failed: 0, Partial: 0, PassRate: 1.0},
	})

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-max-pass-rate-drop", "0.1",
		"-max-failed-increase", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected trend gate failure")
	}
	if !strings.Contains(err.Error(), "pass_rate_drop") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayTrendGatesRequireTrendDir(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-max-pass-rate-drop", "0.1"}, &out); err == nil {
		t.Fatal("expected max-pass-rate-drop validation error")
	}
	if err := run([]string{"-max-failed-increase", "0"}, &out); err == nil {
		t.Fatal("expected max-failed-increase validation error")
	}
}

func TestRunReplaySaveTrend(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-run-id", "abc",
		"-save-trend",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	files, err := filepath.Glob(filepath.Join(history, "*.json"))
	if err != nil {
		t.Fatalf("Glob returned error: %v", err)
	}
	if len(files) != 1 {
		t.Fatalf("expected one saved trend file, got %d", len(files))
	}
	buf, err := os.ReadFile(files[0])
	if err != nil {
		t.Fatalf("ReadFile returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(buf, &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Options.RunID != "abc" {
		t.Fatalf("unexpected run id: %s", rep.Options.RunID)
	}
}

func TestRunReplaySaveTrendCustomName(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-save-trend",
		"-save-trend-name", "custom-report",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	if _, err := os.Stat(filepath.Join(history, "custom-report.json")); err != nil {
		t.Fatalf("expected custom saved report: %v", err)
	}
}

func TestRunReplaySaveTrendRequiresTrendDir(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-save-trend"}, &out); err == nil {
		t.Fatal("expected save-trend validation error")
	}
}

func TestRunReplaySaveTrendPrune(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	if err := os.MkdirAll(history, 0o755); err != nil {
		t.Fatalf("MkdirAll history returned error: %v", err)
	}
	writeTrendReport(t, filepath.Join(history, "r1.json"), replayReport{
		GeneratedAt: "2026-01-01T00:00:00Z",
		Options:     replayOptions{RunID: "h1"},
		Summary:     replaySummary{Processed: 1, Success: 1, PassRate: 1},
	})
	writeTrendReport(t, filepath.Join(history, "r2.json"), replayReport{
		GeneratedAt: "2026-01-02T00:00:00Z",
		Options:     replayOptions{RunID: "h2"},
		Summary:     replaySummary{Processed: 1, Success: 1, PassRate: 1},
	})
	writeTrendReport(t, filepath.Join(history, "r3.json"), replayReport{
		GeneratedAt: "2026-01-03T00:00:00Z",
		Options:     replayOptions{RunID: "h3"},
		Summary:     replaySummary{Processed: 1, Success: 1, PassRate: 1},
	})

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-trend-limit", "2",
		"-run-id", "cur",
		"-save-trend",
		"-save-trend-prune",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	files, err := filepath.Glob(filepath.Join(history, "*.json"))
	if err != nil {
		t.Fatalf("Glob returned error: %v", err)
	}
	if len(files) != 2 {
		t.Fatalf("expected 2 trend files after prune, got %d", len(files))
	}
	foundCur := false
	for _, p := range files {
		buf, err := os.ReadFile(p)
		if err != nil {
			t.Fatalf("ReadFile %s returned error: %v", p, err)
		}
		var rep replayReport
		if err := json.Unmarshal(buf, &rep); err != nil {
			t.Fatalf("Unmarshal %s returned error: %v", p, err)
		}
		if rep.Options.RunID == "cur" {
			foundCur = true
		}
	}
	if !foundCur {
		t.Fatal("expected current saved trend report to be preserved")
	}
}

func TestRunReplaySaveTrendPruneValidation(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-save-trend-prune"}, &out); err == nil {
		t.Fatal("expected save-trend-prune validation error without save-trend")
	}
	if err := run([]string{"-save-trend", "-save-trend-prune", "-trend-dir", ".", "-trend-limit", "0"}, &out); err == nil {
		t.Fatal("expected save-trend-prune validation error for trend-limit")
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

func writeTrendReport(t *testing.T, path string, rep replayReport) {
	t.Helper()
	buf, err := json.MarshalIndent(rep, "", "  ")
	if err != nil {
		t.Fatalf("MarshalIndent returned error: %v", err)
	}
	if err := os.WriteFile(path, buf, 0o644); err != nil {
		t.Fatalf("WriteFile trend report returned error: %v", err)
	}
}
