package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
	"github.com/KairuiKin/go-olespec/pkg/olextract"
)

type replayOptions struct {
	Root            string   `json:"root"`
	Extensions      []string `json:"extensions"`
	Mode            string   `json:"mode"`
	BaselineReport  string   `json:"baseline_report,omitempty"`
	IncludeRaw      bool     `json:"include_raw"`
	DetectImages    bool     `json:"detect_images"`
	DetectOLEDS     bool     `json:"detect_oleds"`
	UnwrapOle10     bool     `json:"unwrap_ole10native"`
	Deduplicate     bool     `json:"deduplicate"`
	MaxDepth        int      `json:"max_depth"`
	MaxArtifacts    int      `json:"max_artifacts"`
	MaxTotalBytes   int64    `json:"max_total_bytes"`
	MaxArtifactSize int64    `json:"max_artifact_size"`
}

type replayFileResult struct {
	Path           string `json:"path"`
	Size           int64  `json:"size"`
	DurationMS     int64  `json:"duration_ms"`
	Success        bool   `json:"success"`
	Partial        bool   `json:"partial"`
	ArtifactsTotal int    `json:"artifacts_total"`
	ArtifactsOK    int    `json:"artifacts_ok"`
	ArtifactsFail  int    `json:"artifacts_failed"`
	Warnings       int    `json:"warnings"`
	ErrorCode      string `json:"error_code,omitempty"`
	ErrorMessage   string `json:"error_message,omitempty"`
}

type replaySummary struct {
	ScannedFiles int     `json:"scanned_files"`
	MatchedFiles int     `json:"matched_files"`
	Processed    int     `json:"processed"`
	Success      int     `json:"success"`
	Failed       int     `json:"failed"`
	Partial      int     `json:"partial"`
	PassRate     float64 `json:"pass_rate"`
	DurationMS   int64   `json:"duration_ms"`
}

type replayReport struct {
	GeneratedAt string             `json:"generated_at"`
	Options     replayOptions      `json:"options"`
	Summary     replaySummary      `json:"summary"`
	Files       []replayFileResult `json:"files"`
	Baseline    *replayBaseline    `json:"baseline,omitempty"`
	Gate        replayGateResult   `json:"gate"`
}

type replayBaseline struct {
	BaselinePath      string             `json:"baseline_path"`
	BaselineGenerated string             `json:"baseline_generated_at,omitempty"`
	BaselineFiles     int                `json:"baseline_files"`
	CurrentFiles      int                `json:"current_files"`
	NewFiles          int                `json:"new_files"`
	RemovedFiles      int                `json:"removed_files"`
	NewlyFailed       int                `json:"newly_failed"`
	Fixed             int                `json:"fixed"`
	ErrorCodeChanged  int                `json:"error_code_changed"`
	NewlyPartial      int                `json:"newly_partial"`
	Regressions       []replayRegression `json:"regressions,omitempty"`
}

type replayRegression struct {
	Path     string `json:"path"`
	Kind     string `json:"kind"`
	Baseline string `json:"baseline"`
	Current  string `json:"current"`
}

type replayGateResult struct {
	Enabled        bool     `json:"enabled"`
	Passed         bool     `json:"passed"`
	MinPassRate    *float64 `json:"min_pass_rate,omitempty"`
	MaxFailed      *int     `json:"max_failed,omitempty"`
	MaxNewlyFailed *int     `json:"max_newly_failed,omitempty"`
	Failures       []string `json:"failures,omitempty"`
}

func main() {
	if err := run(os.Args[1:], os.Stdout); err != nil {
		fmt.Fprintln(os.Stderr, err.Error())
		os.Exit(1)
	}
}

func run(args []string, out io.Writer) error {
	if out == nil {
		return errors.New("output writer is nil")
	}
	fset := flag.NewFlagSet("corpus-replay", flag.ContinueOnError)
	fset.SetOutput(io.Discard)

	var (
		root            = fset.String("root", ".", "root directory for corpus files")
		extCSV          = fset.String("ext", ".doc,.dot,.xls,.xlt,.ppt,.pot,.ole,.cfb", "comma-separated file extensions; empty means all files")
		modeStr         = fset.String("mode", "lenient", "parse mode: strict|lenient")
		baselinePath    = fset.String("baseline", "", "path to baseline replay report JSON for regression diff")
		includeRaw      = fset.Bool("include-raw", false, "include raw artifact payloads in extraction")
		detectImages    = fset.Bool("detect-images", true, "enable image signature detection")
		detectOLEDS     = fset.Bool("detect-oleds", true, "enable OLEDS stream detection")
		unwrapOle10     = fset.Bool("unwrap-ole10native", true, "enable recursive Ole10Native unwrapping")
		dedup           = fset.Bool("deduplicate", true, "enable SHA-256 dedup")
		maxDepth        = fset.Int("max-depth", 16, "max recursive extraction depth")
		maxArtifacts    = fset.Int("max-artifacts", 4096, "max artifacts per file")
		maxTotalBytes   = fset.Int64("max-total-bytes", 64<<20, "max total extracted bytes per file")
		maxArtifactSize = fset.Int64("max-artifact-size", 32<<20, "max single artifact size in bytes")
		minPassRate     = fset.Float64("min-pass-rate", -1, "gate: minimum acceptable pass rate in [0,1], negative disables")
		maxFailed       = fset.Int("max-failed", -1, "gate: maximum allowed failed files, negative disables")
		maxNewlyFailed  = fset.Int("max-newly-failed", -1, "gate: maximum allowed newly failed files vs baseline, negative disables")
		outputPath      = fset.String("output", "", "output report path; empty prints JSON to stdout")
	)
	if err := fset.Parse(args); err != nil {
		return err
	}
	if *maxNewlyFailed >= 0 && strings.TrimSpace(*baselinePath) == "" {
		return errors.New("max-newly-failed requires -baseline")
	}
	if *minPassRate > 1 {
		return errors.New("min-pass-rate must be <= 1")
	}

	mode, err := parseMode(*modeStr)
	if err != nil {
		return err
	}
	extensions := parseExtensions(*extCSV)
	absRoot, err := filepath.Abs(*root)
	if err != nil {
		return err
	}

	start := time.Now()
	scanned := 0
	matched := make([]string, 0, 128)
	err = filepath.WalkDir(absRoot, func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		if d.Type()&os.ModeSymlink != 0 {
			return nil
		}
		scanned++
		if matchesExt(path, extensions) {
			matched = append(matched, path)
		}
		return nil
	})
	if err != nil {
		return err
	}
	sort.Strings(matched)

	opt := replayOptions{
		Root:            absRoot,
		Extensions:      append([]string(nil), extensions...),
		Mode:            strings.ToLower(*modeStr),
		BaselineReport:  strings.TrimSpace(*baselinePath),
		IncludeRaw:      *includeRaw,
		DetectImages:    *detectImages,
		DetectOLEDS:     *detectOLEDS,
		UnwrapOle10:     *unwrapOle10,
		Deduplicate:     *dedup,
		MaxDepth:        *maxDepth,
		MaxArtifacts:    *maxArtifacts,
		MaxTotalBytes:   *maxTotalBytes,
		MaxArtifactSize: *maxArtifactSize,
	}
	report := replayReport{
		GeneratedAt: time.Now().UTC().Format(time.RFC3339Nano),
		Options:     opt,
		Summary: replaySummary{
			ScannedFiles: scanned,
			MatchedFiles: len(matched),
		},
		Files: make([]replayFileResult, 0, len(matched)),
		Gate: replayGateResult{
			Enabled: false,
			Passed:  true,
		},
	}

	for _, p := range matched {
		info, statErr := os.Stat(p)
		size := int64(0)
		if statErr == nil {
			size = info.Size()
		}
		rel := p
		if r, relErr := filepath.Rel(absRoot, p); relErr == nil {
			rel = filepath.ToSlash(r)
		}

		fileStart := time.Now()
		item := replayFileResult{
			Path:       rel,
			Size:       size,
			DurationMS: 0,
		}

		rep, extractErr := olextract.ExtractFile(
			p,
			olecfb.OpenOptions{Mode: mode},
			olecfb.ExtractOptions{
				Mode:              mode,
				IncludeRaw:        *includeRaw,
				DetectImages:      *detectImages,
				DetectOLEDS:       *detectOLEDS,
				UnwrapOle10Native: *unwrapOle10,
				Deduplicate:       *dedup,
				Limits: olecfb.ExtractLimits{
					MaxDepth:        *maxDepth,
					MaxArtifacts:    *maxArtifacts,
					MaxTotalBytes:   *maxTotalBytes,
					MaxArtifactSize: *maxArtifactSize,
				},
			},
		)
		item.DurationMS = time.Since(fileStart).Milliseconds()
		report.Summary.Processed++
		if extractErr != nil {
			item.Success = false
			var oe *olecfb.OLEError
			if errors.As(extractErr, &oe) {
				item.ErrorCode = string(oe.Code)
				item.ErrorMessage = oe.Message
			} else {
				item.ErrorMessage = extractErr.Error()
			}
			report.Summary.Failed++
		} else {
			item.Success = true
			item.Partial = rep.Partial
			item.ArtifactsTotal = rep.Stats.ArtifactsTotal
			item.ArtifactsOK = rep.Stats.ArtifactsOK
			item.ArtifactsFail = rep.Stats.ArtifactsFailed
			item.Warnings = len(rep.Warnings)
			report.Summary.Success++
			if rep.Partial {
				report.Summary.Partial++
			}
		}
		report.Files = append(report.Files, item)
	}

	report.Summary.DurationMS = time.Since(start).Milliseconds()
	if report.Summary.Processed > 0 {
		report.Summary.PassRate = float64(report.Summary.Success) / float64(report.Summary.Processed)
	}
	if strings.TrimSpace(*baselinePath) != "" {
		baseline, loadErr := loadReplayReport(*baselinePath)
		if loadErr != nil {
			return loadErr
		}
		report.Baseline = diffReplayReport(*baselinePath, baseline, report)
	}

	var minPassRatePtr *float64
	if *minPassRate >= 0 {
		v := *minPassRate
		minPassRatePtr = &v
	}
	var maxFailedPtr *int
	if *maxFailed >= 0 {
		v := *maxFailed
		maxFailedPtr = &v
	}
	var maxNewlyFailedPtr *int
	if *maxNewlyFailed >= 0 {
		v := *maxNewlyFailed
		maxNewlyFailedPtr = &v
	}
	gateErr := evaluateGates(&report, minPassRatePtr, maxFailedPtr, maxNewlyFailedPtr)

	buf, err := json.MarshalIndent(report, "", "  ")
	if err != nil {
		return err
	}
	if *outputPath == "" {
		_, err = out.Write(buf)
		if err != nil {
			return err
		}
		_, err = out.Write([]byte("\n"))
		if err != nil {
			return err
		}
		return gateErr
	}

	outAbs, err := filepath.Abs(*outputPath)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(outAbs), 0o755); err != nil {
		return err
	}
	if err := os.WriteFile(outAbs, buf, 0o644); err != nil {
		return err
	}
	return gateErr
}

func loadReplayReport(path string) (replayReport, error) {
	p := strings.TrimSpace(path)
	if p == "" {
		return replayReport{}, errors.New("baseline report path is empty")
	}
	buf, err := os.ReadFile(p)
	if err != nil {
		return replayReport{}, err
	}
	var rep replayReport
	if err := json.Unmarshal(buf, &rep); err != nil {
		return replayReport{}, err
	}
	return rep, nil
}

func diffReplayReport(baselinePath string, base replayReport, cur replayReport) *replayBaseline {
	out := &replayBaseline{
		BaselinePath:      baselinePath,
		BaselineGenerated: base.GeneratedAt,
		BaselineFiles:     len(base.Files),
		CurrentFiles:      len(cur.Files),
		Regressions:       make([]replayRegression, 0, 16),
	}

	baseByPath := map[string]replayFileResult{}
	curByPath := map[string]replayFileResult{}
	for _, f := range base.Files {
		baseByPath[f.Path] = f
	}
	for _, f := range cur.Files {
		curByPath[f.Path] = f
	}

	paths := make([]string, 0, len(curByPath))
	for p := range curByPath {
		paths = append(paths, p)
	}
	sort.Strings(paths)
	for _, p := range paths {
		cf := curByPath[p]
		bf, ok := baseByPath[p]
		if !ok {
			out.NewFiles++
			continue
		}
		if bf.Success && !cf.Success {
			out.NewlyFailed++
			out.Regressions = append(out.Regressions, replayRegression{
				Path:     p,
				Kind:     "new_failure",
				Baseline: fileState(bf),
				Current:  fileState(cf),
			})
		}
		if !bf.Success && cf.Success {
			out.Fixed++
		}
		if !bf.Success && !cf.Success && bf.ErrorCode != cf.ErrorCode {
			out.ErrorCodeChanged++
			out.Regressions = append(out.Regressions, replayRegression{
				Path:     p,
				Kind:     "error_code_changed",
				Baseline: fileState(bf),
				Current:  fileState(cf),
			})
		}
		if !bf.Partial && cf.Partial {
			out.NewlyPartial++
			out.Regressions = append(out.Regressions, replayRegression{
				Path:     p,
				Kind:     "became_partial",
				Baseline: fileState(bf),
				Current:  fileState(cf),
			})
		}
	}
	for p := range baseByPath {
		if _, ok := curByPath[p]; !ok {
			out.RemovedFiles++
		}
	}
	sort.Slice(out.Regressions, func(i, j int) bool {
		if out.Regressions[i].Path == out.Regressions[j].Path {
			return out.Regressions[i].Kind < out.Regressions[j].Kind
		}
		return out.Regressions[i].Path < out.Regressions[j].Path
	})
	if len(out.Regressions) == 0 {
		out.Regressions = nil
	}
	return out
}

func fileState(v replayFileResult) string {
	if v.Success {
		if v.Partial {
			return "success(partial)"
		}
		return "success"
	}
	if v.ErrorCode != "" {
		return "failed:" + v.ErrorCode
	}
	return "failed"
}

func evaluateGates(report *replayReport, minPassRate *float64, maxFailed *int, maxNewlyFailed *int) error {
	if report == nil {
		return errors.New("nil report")
	}
	report.Gate = replayGateResult{
		Enabled:        minPassRate != nil || maxFailed != nil || maxNewlyFailed != nil,
		Passed:         true,
		MinPassRate:    minPassRate,
		MaxFailed:      maxFailed,
		MaxNewlyFailed: maxNewlyFailed,
		Failures:       nil,
	}
	if !report.Gate.Enabled {
		return nil
	}
	if minPassRate != nil && report.Summary.PassRate < *minPassRate {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("pass_rate %.6f < min_pass_rate %.6f", report.Summary.PassRate, *minPassRate))
	}
	if maxFailed != nil && report.Summary.Failed > *maxFailed {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("failed %d > max_failed %d", report.Summary.Failed, *maxFailed))
	}
	if maxNewlyFailed != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_newly_failed gate requires baseline diff")
		} else if report.Baseline.NewlyFailed > *maxNewlyFailed {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("newly_failed %d > max_newly_failed %d", report.Baseline.NewlyFailed, *maxNewlyFailed))
		}
	}
	if len(report.Gate.Failures) == 0 {
		return nil
	}
	report.Gate.Passed = false
	return fmt.Errorf("gate failed: %s", strings.Join(report.Gate.Failures, "; "))
}

func parseMode(v string) (olecfb.ParseMode, error) {
	switch strings.ToLower(strings.TrimSpace(v)) {
	case "", "lenient":
		return olecfb.Lenient, nil
	case "strict":
		return olecfb.Strict, nil
	default:
		return olecfb.Lenient, fmt.Errorf("invalid mode %q, expected strict|lenient", v)
	}
}

func parseExtensions(csv string) []string {
	csv = strings.TrimSpace(csv)
	if csv == "" {
		return nil
	}
	parts := strings.Split(csv, ",")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.ToLower(strings.TrimSpace(p))
		if p == "" {
			continue
		}
		if !strings.HasPrefix(p, ".") {
			p = "." + p
		}
		out = append(out, p)
	}
	return out
}

func matchesExt(path string, exts []string) bool {
	if len(exts) == 0 {
		return true
	}
	ext := strings.ToLower(filepath.Ext(path))
	for _, e := range exts {
		if ext == e {
			return true
		}
	}
	return false
}
