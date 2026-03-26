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

func probeFirstReplayErrorCode(t *testing.T, root string) string {
	t.Helper()
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb"}, &out); err != nil {
		t.Fatalf("probe run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal probe returned error: %v", err)
	}
	for k := range rep.Summary.ErrorCodes {
		return k
	}
	t.Fatal("expected at least one error code")
	return ""
}

func TestRunReplayReportFilesFailed(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-report-files", "failed"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one failed file entry, got %d", len(rep.Files))
	}
	if rep.Files[0].Success {
		t.Fatal("expected failed-only entries")
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportFilesSuccess(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-report-files", "success"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one success file entry, got %d", len(rep.Files))
	}
	if !rep.Files[0].Success {
		t.Fatal("expected success-only entries")
	}
	if rep.Files[0].Path != "ok.cfb" {
		t.Fatalf("unexpected success file path: %s", rep.Files[0].Path)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportFilesNone(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-report-files", "none"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 0 {
		t.Fatalf("expected no file entries, got %d", len(rep.Files))
	}
	if rep.Summary.OmittedFiles != rep.Summary.Processed {
		t.Fatalf("expected all processed files omitted, summary=%+v", rep.Summary)
	}
}

func TestRunReplayReportFilesIssues(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-report-files", "issues"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one issue file entry, got %d", len(rep.Files))
	}
	if rep.Files[0].Success {
		t.Fatal("expected issue file to be failed")
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportFilesIssuesIncludePartial(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-report-files", "issues",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one issue file entry, got %d", len(rep.Files))
	}
	if !rep.Files[0].Success || !rep.Files[0].Partial {
		t.Fatalf("expected partial success entry, got %+v", rep.Files[0])
	}
}

func TestRunReplayReportFilesPartial(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "partial.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile partial returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-report-files", "partial",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one partial file entry, got %d", len(rep.Files))
	}
	if !rep.Files[0].Success || !rep.Files[0].Partial {
		t.Fatalf("expected partial-success file entry, got %+v", rep.Files[0])
	}
	if rep.Files[0].Path != "partial.cfb" {
		t.Fatalf("unexpected partial file path: %s", rep.Files[0].Path)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportFilesSuccessIncludePartial(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "partial.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile partial returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-report-files", "success",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one success file entry, got %d", len(rep.Files))
	}
	if !rep.Files[0].Success || !rep.Files[0].Partial {
		t.Fatalf("expected partial success file entry, got %+v", rep.Files[0])
	}
}

func TestRunReplayReportFilesWarnings(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "warn.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile warn returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-report-files", "warnings",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one warning file entry, got %d", len(rep.Files))
	}
	if rep.Files[0].Warnings == 0 {
		t.Fatalf("expected warning entry, got %+v", rep.Files[0])
	}
	if rep.Files[0].Path != "warn.cfb" {
		t.Fatalf("expected warn.cfb entry, got %s", rep.Files[0].Path)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportFilesClean(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-report-files", "clean",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one clean file entry, got %d", len(rep.Files))
	}
	if !rep.Files[0].Success || rep.Files[0].Partial || rep.Files[0].Warnings != 0 {
		t.Fatalf("expected clean file entry, got %+v", rep.Files[0])
	}
	if rep.Files[0].Path != "ok.cfb" {
		t.Fatalf("unexpected clean file path: %s", rep.Files[0].Path)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportFilesCleanExcludePartial(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "partial.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile partial returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-report-files", "clean",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 0 {
		t.Fatalf("expected no clean file entries, got %d", len(rep.Files))
	}
	if rep.Summary.ReportedFiles != 0 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportLimit(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-report-limit", "1"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one file entry with report-limit=1, got %d", len(rep.Files))
	}
	if rep.Summary.Processed != 2 {
		t.Fatalf("unexpected processed: %d", rep.Summary.Processed)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
	if rep.Options.ReportLimit != 1 {
		t.Fatalf("unexpected report limit option: %d", rep.Options.ReportLimit)
	}
}

func TestRunReplayReportLimitWithFailedPolicy(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-report-files", "failed",
		"-report-limit", "0",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 0 {
		t.Fatalf("expected no file entries, got %d", len(rep.Files))
	}
	if rep.Summary.ReportedFiles != 0 || rep.Summary.OmittedFiles != 2 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestRunReplayReportOffsetAndLimit(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "b.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile b returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "c.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile c returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-report-sort", "path",
		"-report-offset", "1",
		"-report-limit", "1",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 3 {
		t.Fatalf("unexpected processed: %d", rep.Summary.Processed)
	}
	if len(rep.Files) != 1 || rep.Files[0].Path != "b.cfb" {
		t.Fatalf("unexpected paged files: %+v", rep.Files)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 2 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
	if rep.Options.ReportOffset != 1 || rep.Options.ReportLimit != 1 {
		t.Fatalf("unexpected pagination options: %+v", rep.Options)
	}
}

func TestRunReplayReportErrorCodesInclude(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	code := probeFirstReplayErrorCode(t, root)

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-report-files", "all",
		"-report-error-codes", strings.ToLower(code),
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one included file entry, got %d", len(rep.Files))
	}
	if rep.Files[0].Path != "bad.cfb" {
		t.Fatalf("expected bad.cfb entry, got %s", rep.Files[0].Path)
	}
	if rep.Summary.Processed != 2 {
		t.Fatalf("unexpected processed: %d", rep.Summary.Processed)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
	if len(rep.Options.ReportErrorCodes) != 1 || rep.Options.ReportErrorCodes[0] != code {
		t.Fatalf("unexpected report error code options: %+v", rep.Options.ReportErrorCodes)
	}
}

func TestRunReplayReportErrorCodesExclude(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	code := probeFirstReplayErrorCode(t, root)

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-report-files", "all",
		"-report-exclude-error-codes", code,
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 1 {
		t.Fatalf("expected one non-excluded file entry, got %d", len(rep.Files))
	}
	if rep.Files[0].Path != "ok.cfb" {
		t.Fatalf("expected ok.cfb entry, got %s", rep.Files[0].Path)
	}
	if rep.Summary.Processed != 2 {
		t.Fatalf("unexpected processed: %d", rep.Summary.Processed)
	}
	if rep.Summary.ReportedFiles != 1 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
	if len(rep.Options.ReportExcludeErrorCodes) != 1 || rep.Options.ReportExcludeErrorCodes[0] != code {
		t.Fatalf("unexpected report exclude error code options: %+v", rep.Options.ReportExcludeErrorCodes)
	}
}

func TestRunReplayReportErrorCodesExcludeWins(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "bad.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile bad returned error: %v", err)
	}
	code := probeFirstReplayErrorCode(t, root)

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-report-files", "failed",
		"-report-error-codes", code,
		"-report-exclude-error-codes", code,
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if len(rep.Files) != 0 {
		t.Fatalf("expected exclude to win and remove all entries, got %d", len(rep.Files))
	}
	if rep.Summary.ReportedFiles != 0 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestApplyReportFilePolicySortDurationDesc(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "a.cfb", DurationMS: 10, Success: true},
			{Path: "c.cfb", DurationMS: 30, Success: true},
			{Path: "b.cfb", DurationMS: 20, Success: true},
		},
	}
	applyReportFilePolicy(rep, "all", "duration-desc", 0, 2, nil, nil)
	if len(rep.Files) != 2 {
		t.Fatalf("expected 2 file entries, got %d", len(rep.Files))
	}
	if rep.Files[0].Path != "c.cfb" || rep.Files[1].Path != "b.cfb" {
		t.Fatalf("unexpected duration-desc order: %+v", rep.Files)
	}
	if rep.Summary.ReportedFiles != 2 || rep.Summary.OmittedFiles != 1 {
		t.Fatalf("unexpected reported/omitted: %+v", rep.Summary)
	}
}

func TestApplyReportFilePolicySortFailedFirst(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "ok.cfb", Success: true, Partial: false, Warnings: 0, DurationMS: 5},
			{Path: "partial.cfb", Success: true, Partial: true, Warnings: 1, DurationMS: 4},
			{Path: "bad1.cfb", Success: false, Partial: false, Warnings: 0, DurationMS: 2},
			{Path: "bad2.cfb", Success: false, Partial: false, Warnings: 0, DurationMS: 8},
		},
	}
	applyReportFilePolicy(rep, "all", "failed-first", 0, -1, nil, nil)
	if len(rep.Files) != 4 {
		t.Fatalf("expected 4 file entries, got %d", len(rep.Files))
	}
	got := []string{rep.Files[0].Path, rep.Files[1].Path, rep.Files[2].Path, rep.Files[3].Path}
	want := []string{"bad2.cfb", "bad1.cfb", "partial.cfb", "ok.cfb"}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("unexpected failed-first order: got=%v want=%v", got, want)
		}
	}
}

func TestApplyReportFilePolicySortSizeDesc(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "a.cfb", Size: 20, Success: true},
			{Path: "c.cfb", Size: 10, Success: true},
			{Path: "b.cfb", Size: 20, Success: true},
		},
	}
	applyReportFilePolicy(rep, "all", "size-desc", 0, -1, nil, nil)
	if len(rep.Files) != 3 {
		t.Fatalf("expected 3 file entries, got %d", len(rep.Files))
	}
	got := []string{rep.Files[0].Path, rep.Files[1].Path, rep.Files[2].Path}
	want := []string{"a.cfb", "b.cfb", "c.cfb"}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("unexpected size-desc order: got=%v want=%v", got, want)
		}
	}
}

func TestApplyReportFilePolicySortArtifactsDesc(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "a.cfb", ArtifactsTotal: 2, Success: true},
			{Path: "c.cfb", ArtifactsTotal: 1, Success: true},
			{Path: "b.cfb", ArtifactsTotal: 2, Success: true},
		},
	}
	applyReportFilePolicy(rep, "all", "artifacts-desc", 0, -1, nil, nil)
	if len(rep.Files) != 3 {
		t.Fatalf("expected 3 file entries, got %d", len(rep.Files))
	}
	got := []string{rep.Files[0].Path, rep.Files[1].Path, rep.Files[2].Path}
	want := []string{"a.cfb", "b.cfb", "c.cfb"}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("unexpected artifacts-desc order: got=%v want=%v", got, want)
		}
	}
}

func TestApplyReportFilePolicySortArtifactsFailedDesc(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "a.cfb", ArtifactsFail: 2, Success: true},
			{Path: "c.cfb", ArtifactsFail: 1, Success: true},
			{Path: "b.cfb", ArtifactsFail: 2, Success: true},
		},
	}
	applyReportFilePolicy(rep, "all", "artifacts-failed-desc", 0, -1, nil, nil)
	if len(rep.Files) != 3 {
		t.Fatalf("expected 3 file entries, got %d", len(rep.Files))
	}
	got := []string{rep.Files[0].Path, rep.Files[1].Path, rep.Files[2].Path}
	want := []string{"a.cfb", "b.cfb", "c.cfb"}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("unexpected artifacts-failed-desc order: got=%v want=%v", got, want)
		}
	}
}

func TestApplyReportFilePolicySortWarningsDesc(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "a.cfb", Warnings: 2, Success: true},
			{Path: "c.cfb", Warnings: 1, Success: true},
			{Path: "b.cfb", Warnings: 2, Success: true},
		},
	}
	applyReportFilePolicy(rep, "all", "warnings-desc", 0, -1, nil, nil)
	if len(rep.Files) != 3 {
		t.Fatalf("expected 3 file entries, got %d", len(rep.Files))
	}
	got := []string{rep.Files[0].Path, rep.Files[1].Path, rep.Files[2].Path}
	want := []string{"a.cfb", "b.cfb", "c.cfb"}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("unexpected warnings-desc order: got=%v want=%v", got, want)
		}
	}
}

func TestApplyReportFilePolicySortErrorCode(t *testing.T) {
	rep := &replayReport{
		Files: []replayFileResult{
			{Path: "ok.cfb", ErrorCode: "", Success: true},
			{Path: "z1.cfb", ErrorCode: "ZZZ", Success: false},
			{Path: "a2.cfb", ErrorCode: "aaa", Success: false},
			{Path: "a1.cfb", ErrorCode: "AAA", Success: false},
		},
	}
	applyReportFilePolicy(rep, "all", "error-code", 0, -1, nil, nil)
	if len(rep.Files) != 4 {
		t.Fatalf("expected 4 file entries, got %d", len(rep.Files))
	}
	got := []string{rep.Files[0].Path, rep.Files[1].Path, rep.Files[2].Path, rep.Files[3].Path}
	want := []string{"a1.cfb", "a2.cfb", "z1.cfb", "ok.cfb"}
	for i := range want {
		if got[i] != want[i] {
			t.Fatalf("unexpected error-code order: got=%v want=%v", got, want)
		}
	}
}

func TestRunReplayReportFilesInvalid(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-report-files", "bad"}, &out); err == nil {
		t.Fatal("expected report-files validation error")
	}
	if err := run([]string{"-report-sort", "bad"}, &out); err == nil {
		t.Fatal("expected report-sort validation error")
	}
	if err := run([]string{"-report-offset", "-1"}, &out); err == nil {
		t.Fatal("expected report-offset validation error")
	}
	if err := run([]string{"-report-limit", "-2"}, &out); err == nil {
		t.Fatal("expected report-limit validation error")
	}
}

func TestRunReplayIncludeGlob(t *testing.T) {
	root := t.TempDir()
	keepDir := filepath.Join(root, "keep")
	skipDir := filepath.Join(root, "skip")
	if err := os.MkdirAll(keepDir, 0o755); err != nil {
		t.Fatalf("MkdirAll keep returned error: %v", err)
	}
	if err := os.MkdirAll(skipDir, 0o755); err != nil {
		t.Fatalf("MkdirAll skip returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(keepDir, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile keep returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(skipDir, "b.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile skip returned error: %v", err)
	}

	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-include-glob", "keep/*.cfb"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 1 {
		t.Fatalf("expected 1 processed file, got %d", rep.Summary.Processed)
	}
	if rep.Summary.Success != 1 || rep.Summary.Failed != 0 {
		t.Fatalf("unexpected summary: %+v", rep.Summary)
	}
	if len(rep.Options.IncludeGlobs) != 1 || rep.Options.IncludeGlobs[0] != "keep/*.cfb" {
		t.Fatalf("unexpected include globs: %+v", rep.Options.IncludeGlobs)
	}
}

func TestRunReplayExcludeGlob(t *testing.T) {
	root := t.TempDir()
	keepDir := filepath.Join(root, "keep")
	skipDir := filepath.Join(root, "skip")
	if err := os.MkdirAll(keepDir, 0o755); err != nil {
		t.Fatalf("MkdirAll keep returned error: %v", err)
	}
	if err := os.MkdirAll(skipDir, 0o755); err != nil {
		t.Fatalf("MkdirAll skip returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(keepDir, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile keep returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(skipDir, "b.cfb"), []byte("bad"), 0o644); err != nil {
		t.Fatalf("WriteFile skip returned error: %v", err)
	}

	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-exclude-glob", "skip/*.cfb"}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 1 {
		t.Fatalf("expected 1 processed file, got %d", rep.Summary.Processed)
	}
	if rep.Summary.Success != 1 || rep.Summary.Failed != 0 {
		t.Fatalf("unexpected summary: %+v", rep.Summary)
	}
	if len(rep.Options.ExcludeGlobs) != 1 || rep.Options.ExcludeGlobs[0] != "skip/*.cfb" {
		t.Fatalf("unexpected exclude globs: %+v", rep.Options.ExcludeGlobs)
	}
}

func TestRunReplayIncludeExcludeGlobPrecedence(t *testing.T) {
	root := t.TempDir()
	keepDir := filepath.Join(root, "keep")
	otherDir := filepath.Join(root, "other")
	if err := os.MkdirAll(keepDir, 0o755); err != nil {
		t.Fatalf("MkdirAll keep returned error: %v", err)
	}
	if err := os.MkdirAll(otherDir, 0o755); err != nil {
		t.Fatalf("MkdirAll other returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(keepDir, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile keep returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(otherDir, "b.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile other returned error: %v", err)
	}

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-include-glob", "*/*.cfb",
		"-exclude-glob", "keep/*.cfb",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 1 {
		t.Fatalf("expected 1 processed file, got %d", rep.Summary.Processed)
	}
	if len(rep.Files) != 1 || rep.Files[0].Path != "other/b.cfb" {
		t.Fatalf("unexpected files: %+v", rep.Files)
	}
}

func TestRunReplayGlobPatternValidation(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-include-glob", "[bad"}, &out); err == nil {
		t.Fatal("expected include-glob validation error")
	}
	if err := run([]string{"-exclude-glob", "[bad"}, &out); err == nil {
		t.Fatal("expected exclude-glob validation error")
	}
}

func TestRunReplayMinFileSizeFilter(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "small.cfb"), []byte("x"), 0o644); err != nil {
		t.Fatalf("WriteFile small returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "big.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile big returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-min-file-size-bytes", "2",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 1 {
		t.Fatalf("expected processed=1, got %d", rep.Summary.Processed)
	}
	if len(rep.Files) != 1 || rep.Files[0].Path != "big.cfb" {
		t.Fatalf("unexpected files: %+v", rep.Files)
	}
}

func TestRunReplayMaxFileSizeFilter(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "small.cfb"), []byte("x"), 0o644); err != nil {
		t.Fatalf("WriteFile small returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "big.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile big returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-file-size-bytes", "10",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 1 {
		t.Fatalf("expected processed=1, got %d", rep.Summary.Processed)
	}
	if len(rep.Files) != 1 || rep.Files[0].Path != "small.cfb" {
		t.Fatalf("unexpected files: %+v", rep.Files)
	}
}

func TestRunReplayFileSizeFilterValidation(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-min-file-size-bytes", "10", "-max-file-size-bytes", "1"}, &out); err == nil {
		t.Fatal("expected min/max file-size validation error")
	}
}

func TestRunReplayMaxMatchedFiles(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "b.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile b returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "c.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile c returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-matched-files", "2",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}
	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.Processed != 2 || rep.Summary.MatchedFiles != 2 {
		t.Fatalf("unexpected processed/matched summary: %+v", rep.Summary)
	}
	if rep.Summary.MatchedFilesTotal != 3 {
		t.Fatalf("unexpected matched_files_total: %d", rep.Summary.MatchedFilesTotal)
	}
	if rep.Summary.TruncatedMatches != 1 {
		t.Fatalf("unexpected truncated_matches: %d", rep.Summary.TruncatedMatches)
	}
	if len(rep.Files) != 2 {
		t.Fatalf("expected 2 files, got %d", len(rep.Files))
	}
	if rep.Files[0].Path != "a.cfb" || rep.Files[1].Path != "b.cfb" {
		t.Fatalf("unexpected replay selection order: %+v", rep.Files)
	}
	if rep.Options.MaxMatchedFiles != 2 {
		t.Fatalf("unexpected max_matched_files option: %d", rep.Options.MaxMatchedFiles)
	}
}

func TestRunReplayMaxMatchedFilesValidation(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-max-matched-files", "-2"}, &out); err == nil {
		t.Fatal("expected max-matched-files validation error")
	}
	if err := run([]string{"-min-scanned-files", "-2"}, &out); err == nil {
		t.Fatal("expected min-scanned-files validation error")
	}
	if err := run([]string{"-max-scanned-files", "-2"}, &out); err == nil {
		t.Fatal("expected max-scanned-files validation error")
	}
	if err := run([]string{"-min-matched-files", "-2"}, &out); err == nil {
		t.Fatal("expected min-matched-files validation error")
	}
	if err := run([]string{"-max-matched-files-total", "-2"}, &out); err == nil {
		t.Fatal("expected max-matched-files-total validation error")
	}
	if err := run([]string{"-max-truncated-matches", "-2"}, &out); err == nil {
		t.Fatal("expected max-truncated-matches validation error")
	}
}

func TestRunReplayMinMatchedFilesGate(t *testing.T) {
	root := t.TempDir()
	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-min-matched-files", "1"}, &out)
	if err == nil {
		t.Fatal("expected min-matched-files gate error")
	}
	if !strings.Contains(err.Error(), "min_matched_files") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.MatchedFiles != 0 || rep.Summary.TruncatedMatches != 0 {
		t.Fatalf("unexpected match summary: %+v", rep.Summary)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMinScannedFilesGate(t *testing.T) {
	root := t.TempDir()
	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-min-scanned-files", "1"}, &out)
	if err == nil {
		t.Fatal("expected min-scanned-files gate error")
	}
	if !strings.Contains(err.Error(), "min_scanned_files") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.ScannedFiles != 0 {
		t.Fatalf("unexpected scanned files: %d", rep.Summary.ScannedFiles)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxScannedFilesGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-max-scanned-files", "0"}, &out)
	if err == nil {
		t.Fatal("expected max-scanned-files gate error")
	}
	if !strings.Contains(err.Error(), "max_scanned_files") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.ScannedFiles != 1 {
		t.Fatalf("unexpected scanned files: %d", rep.Summary.ScannedFiles)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxTruncatedMatchesGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "b.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile b returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-matched-files", "1",
		"-max-truncated-matches", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-truncated-matches gate error")
	}
	if !strings.Contains(err.Error(), "max_truncated_matches") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.TruncatedMatches != 1 {
		t.Fatalf("unexpected truncated matches: %d", rep.Summary.TruncatedMatches)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxMatchedFilesTotalGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "a.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile a returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "b.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile b returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-matched-files-total", "1",
	}, &out)
	if err == nil {
		t.Fatal("expected max-matched-files-total gate error")
	}
	if !strings.Contains(err.Error(), "max_matched_files_total") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.MatchedFilesTotal != 2 {
		t.Fatalf("unexpected matched_files_total: %d", rep.Summary.MatchedFilesTotal)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayFilterCounters(t *testing.T) {
	root := t.TempDir()
	if err := os.MkdirAll(filepath.Join(root, "keep"), 0o755); err != nil {
		t.Fatalf("MkdirAll keep returned error: %v", err)
	}
	if err := os.MkdirAll(filepath.Join(root, "drop"), 0o755); err != nil {
		t.Fatalf("MkdirAll drop returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "a.txt"), []byte("txt"), 0o644); err != nil {
		t.Fatalf("WriteFile a.txt returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "drop", "b.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile drop/b.cfb returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "keep", "small.cfb"), []byte("x"), 0o644); err != nil {
		t.Fatalf("WriteFile keep/small.cfb returned error: %v", err)
	}
	if err := os.WriteFile(filepath.Join(root, "keep", "big.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile keep/big.cfb returned error: %v", err)
	}

	var out bytes.Buffer
	if err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-include-glob", "keep/*.cfb",
		"-min-file-size-bytes", "2",
	}, &out); err != nil {
		t.Fatalf("run returned error: %v", err)
	}

	var rep replayReport
	if err := json.Unmarshal(out.Bytes(), &rep); err != nil {
		t.Fatalf("Unmarshal returned error: %v", err)
	}
	if rep.Summary.ScannedFiles != 4 {
		t.Fatalf("unexpected scanned_files: %d", rep.Summary.ScannedFiles)
	}
	if rep.Summary.MatchedFiles != 1 || rep.Summary.Processed != 1 {
		t.Fatalf("unexpected matched/processed summary: %+v", rep.Summary)
	}
	if rep.Summary.FilteredByExt != 1 {
		t.Fatalf("unexpected filtered_by_ext: %d", rep.Summary.FilteredByExt)
	}
	if rep.Summary.FilteredByPath != 1 {
		t.Fatalf("unexpected filtered_by_path: %d", rep.Summary.FilteredByPath)
	}
	if rep.Summary.FilteredBySize != 1 {
		t.Fatalf("unexpected filtered_by_size: %d", rep.Summary.FilteredBySize)
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

func TestRunReplayBaselineLatest(t *testing.T) {
	root := t.TempDir()
	target := filepath.Join(root, "target.cfb")
	if err := os.WriteFile(target, buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile target returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	if err := os.MkdirAll(history, 0o755); err != nil {
		t.Fatalf("MkdirAll history returned error: %v", err)
	}
	baselinePath := filepath.Join(history, "base.json")
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
		"-trend-dir", history,
		"-baseline-latest",
		"-max-newly-failed", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected gate failure")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Baseline == nil {
		t.Fatal("expected baseline diff")
	}
	if rep.Options.BaselineReport == "" {
		t.Fatal("expected resolved baseline path")
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

func TestRunReplayMinProcessedGate(t *testing.T) {
	root := t.TempDir()
	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-min-processed", "1"}, &out)
	if err == nil {
		t.Fatal("expected min-processed gate error")
	}
	if !strings.Contains(err.Error(), "min_processed") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.Processed != 0 {
		t.Fatalf("expected processed=0, got %d", rep.Summary.Processed)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxProcessedGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{"-root", root, "-ext", ".cfb", "-max-processed", "0"}, &out)
	if err == nil {
		t.Fatal("expected max-processed gate error")
	}
	if !strings.Contains(err.Error(), "max_processed") {
		t.Fatalf("unexpected gate error: %v", err)
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.Processed != 1 {
		t.Fatalf("expected processed=1, got %d", rep.Summary.Processed)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxPartialGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-max-partial", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-partial gate error")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.Partial == 0 {
		t.Fatalf("expected partial files, got %d", rep.Summary.Partial)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxWarningsGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-max-warnings", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-warnings gate error")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.WarningsTotal == 0 {
		t.Fatalf("expected warning count > 0, got %d", rep.Summary.WarningsTotal)
	}
	if rep.Gate.Passed {
		t.Fatal("expected gate fail")
	}
}

func TestRunReplayMaxWarningFilesGate(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-max-artifact-size", "1",
		"-max-warning-files", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected max-warning-files gate error")
	}
	var rep replayReport
	if jsonErr := json.Unmarshal(out.Bytes(), &rep); jsonErr != nil {
		t.Fatalf("Unmarshal returned error: %v", jsonErr)
	}
	if rep.Summary.WarningFiles == 0 {
		t.Fatalf("expected warning files > 0, got %d", rep.Summary.WarningFiles)
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

func TestRunReplayBaselineLatestValidation(t *testing.T) {
	var out bytes.Buffer
	if err := run([]string{"-baseline-latest"}, &out); err == nil {
		t.Fatal("expected baseline-latest validation error without trend-dir")
	}
	if err := run([]string{"-baseline-latest", "-baseline", "a.json", "-trend-dir", "."}, &out); err == nil {
		t.Fatal("expected baseline-latest conflict validation error")
	}
}

func TestRunReplayBaselineLatestNoReports(t *testing.T) {
	root := t.TempDir()
	if err := os.WriteFile(filepath.Join(root, "ok.cfb"), buildSampleCFB(t), 0o644); err != nil {
		t.Fatalf("WriteFile ok returned error: %v", err)
	}
	history := filepath.Join(root, "history")
	if err := os.MkdirAll(history, 0o755); err != nil {
		t.Fatalf("MkdirAll history returned error: %v", err)
	}
	var out bytes.Buffer
	if err := run([]string{"-root", root, "-ext", ".cfb", "-trend-dir", history, "-baseline-latest"}, &out); err == nil {
		t.Fatal("expected baseline-latest no report error")
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
	if err := run([]string{"-max-processed-increase", "0"}, &out); err == nil {
		t.Fatal("expected max-processed-increase validation error")
	}
	if err := run([]string{"-max-processed-drop", "0"}, &out); err == nil {
		t.Fatal("expected max-processed-drop validation error")
	}
	if err := run([]string{"-max-failed-increase", "0"}, &out); err == nil {
		t.Fatal("expected max-failed-increase validation error")
	}
	if err := run([]string{"-max-partial-increase", "0"}, &out); err == nil {
		t.Fatal("expected max-partial-increase validation error")
	}
	if err := run([]string{"-max-warning-increase", "0"}, &out); err == nil {
		t.Fatal("expected max-warning-increase validation error")
	}
}

func TestRunReplayTrendPartialIncreaseGate(t *testing.T) {
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
		Summary:     replaySummary{Processed: 1, Success: 1, Failed: 0, Partial: 0, PassRate: 1.0},
	})

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-max-artifact-size", "1",
		"-max-partial-increase", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected trend partial-increase gate failure")
	}
	if !strings.Contains(err.Error(), "partial_increase") {
		t.Fatalf("unexpected gate error: %v", err)
	}
}

func TestRunReplayTrendWarningIncreaseGate(t *testing.T) {
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
		Summary:     replaySummary{Processed: 1, Success: 1, Failed: 0, Partial: 0, WarningsTotal: 0, PassRate: 1.0},
	})

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-max-artifact-size", "1",
		"-max-warning-increase", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected trend warning-increase gate failure")
	}
	if !strings.Contains(err.Error(), "warning_increase") {
		t.Fatalf("unexpected gate error: %v", err)
	}
}

func TestRunReplayTrendProcessedIncreaseGate(t *testing.T) {
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
		Summary:     replaySummary{Processed: 0, Success: 0, Failed: 0, Partial: 0, WarningsTotal: 0, PassRate: 0.0},
	})

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-max-processed-increase", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected trend processed-increase gate failure")
	}
	if !strings.Contains(err.Error(), "processed_increase") {
		t.Fatalf("unexpected gate error: %v", err)
	}
}

func TestRunReplayTrendProcessedDropGate(t *testing.T) {
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
		Summary:     replaySummary{Processed: 2, Success: 2, Failed: 0, Partial: 0, WarningsTotal: 0, PassRate: 1.0},
	})

	var out bytes.Buffer
	err := run([]string{
		"-root", root,
		"-ext", ".cfb",
		"-trend-dir", history,
		"-max-processed-drop", "0",
	}, &out)
	if err == nil {
		t.Fatal("expected trend processed-drop gate failure")
	}
	if !strings.Contains(err.Error(), "processed_drop") {
		t.Fatalf("unexpected gate error: %v", err)
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
