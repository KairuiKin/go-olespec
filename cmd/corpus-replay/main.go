package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path"
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
	IncludeGlobs    []string `json:"include_globs,omitempty"`
	ExcludeGlobs    []string `json:"exclude_globs,omitempty"`
	Mode            string   `json:"mode"`
	ReportFiles     string   `json:"report_files"`
	RunID           string   `json:"run_id,omitempty"`
	BaselineReport  string   `json:"baseline_report,omitempty"`
	BaselineLatest  bool     `json:"baseline_latest,omitempty"`
	TrendDir        string   `json:"trend_dir,omitempty"`
	TrendLimit      int      `json:"trend_limit,omitempty"`
	SaveTrend       bool     `json:"save_trend,omitempty"`
	SaveTrendName   string   `json:"save_trend_name,omitempty"`
	SaveTrendPrune  bool     `json:"save_trend_prune,omitempty"`
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
	ScannedFiles  int            `json:"scanned_files"`
	MatchedFiles  int            `json:"matched_files"`
	Processed     int            `json:"processed"`
	Success       int            `json:"success"`
	Failed        int            `json:"failed"`
	Partial       int            `json:"partial"`
	WarningsTotal int            `json:"warnings_total"`
	ReportedFiles int            `json:"reported_files"`
	OmittedFiles  int            `json:"omitted_files"`
	PassRate      float64        `json:"pass_rate"`
	DurationMS    int64          `json:"duration_ms"`
	ErrorCodes    map[string]int `json:"error_codes,omitempty"`
}

type replayReport struct {
	GeneratedAt string             `json:"generated_at"`
	Options     replayOptions      `json:"options"`
	Summary     replaySummary      `json:"summary"`
	Files       []replayFileResult `json:"files"`
	Baseline    *replayBaseline    `json:"baseline,omitempty"`
	Trend       *replayTrend       `json:"trend,omitempty"`
	Gate        replayGateResult   `json:"gate"`
}

type replayBaseline struct {
	BaselinePath        string             `json:"baseline_path"`
	BaselineGenerated   string             `json:"baseline_generated_at,omitempty"`
	BaselineFiles       int                `json:"baseline_files"`
	CurrentFiles        int                `json:"current_files"`
	NewFiles            int                `json:"new_files"`
	RemovedFiles        int                `json:"removed_files"`
	NewlyFailed         int                `json:"newly_failed"`
	Fixed               int                `json:"fixed"`
	ErrorCodeChanged    int                `json:"error_code_changed"`
	NewlyPartial        int                `json:"newly_partial"`
	ErrorCodeDelta      map[string]int     `json:"error_code_delta,omitempty"`
	NewErrorCodes       []string           `json:"new_error_codes,omitempty"`
	IncreasedErrorCodes []replayCodeDelta  `json:"increased_error_codes,omitempty"`
	Regressions         []replayRegression `json:"regressions,omitempty"`
}

type replayCodeDelta struct {
	Code     string `json:"code"`
	Baseline int    `json:"baseline"`
	Current  int    `json:"current"`
	Delta    int    `json:"delta"`
}

type replayRegression struct {
	Path     string `json:"path"`
	Kind     string `json:"kind"`
	Baseline string `json:"baseline"`
	Current  string `json:"current"`
}

type replayTrend struct {
	SourceDir   string             `json:"source_dir"`
	Points      []replayTrendPoint `json:"points"`
	LatestDelta *replayTrendDelta  `json:"latest_delta,omitempty"`
}

type replayTrendPoint struct {
	GeneratedAt string  `json:"generated_at"`
	RunID       string  `json:"run_id,omitempty"`
	ReportPath  string  `json:"report_path,omitempty"`
	Processed   int     `json:"processed"`
	Success     int     `json:"success"`
	Failed      int     `json:"failed"`
	Partial     int     `json:"partial"`
	Warnings    int     `json:"warnings"`
	PassRate    float64 `json:"pass_rate"`
}

type replayTrendDelta struct {
	FromGeneratedAt string  `json:"from_generated_at"`
	ToGeneratedAt   string  `json:"to_generated_at"`
	DeltaPassRate   float64 `json:"delta_pass_rate"`
	DeltaFailed     int     `json:"delta_failed"`
	DeltaPartial    int     `json:"delta_partial"`
	DeltaWarnings   int     `json:"delta_warnings"`
}

type replayGateResult struct {
	Enabled                 bool     `json:"enabled"`
	Passed                  bool     `json:"passed"`
	MinProcessed            *int     `json:"min_processed,omitempty"`
	MaxProcessed            *int     `json:"max_processed,omitempty"`
	MinPassRate             *float64 `json:"min_pass_rate,omitempty"`
	MaxFailed               *int     `json:"max_failed,omitempty"`
	MaxPartial              *int     `json:"max_partial,omitempty"`
	MaxWarnings             *int     `json:"max_warnings,omitempty"`
	MaxNewlyFailed          *int     `json:"max_newly_failed,omitempty"`
	MaxNewFiles             *int     `json:"max_new_files,omitempty"`
	MaxRemovedFiles         *int     `json:"max_removed_files,omitempty"`
	MaxNewlyPartial         *int     `json:"max_newly_partial,omitempty"`
	MaxPassRateDrop         *float64 `json:"max_pass_rate_drop,omitempty"`
	MaxFailedIncrease       *int     `json:"max_failed_increase,omitempty"`
	MaxPartialIncrease      *int     `json:"max_partial_increase,omitempty"`
	MaxWarningIncrease      *int     `json:"max_warning_increase,omitempty"`
	DenyErrorCodes          []string `json:"deny_error_codes,omitempty"`
	MaxNewErrorCodes        *int     `json:"max_new_error_codes,omitempty"`
	MaxErrorCodeRegressions *int     `json:"max_error_code_regressions,omitempty"`
	Failures                []string `json:"failures,omitempty"`
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
		root                    = fset.String("root", ".", "root directory for corpus files")
		extCSV                  = fset.String("ext", ".doc,.dot,.xls,.xlt,.ppt,.pot,.ole,.cfb", "comma-separated file extensions; empty means all files")
		includeGlobCSV          = fset.String("include-glob", "", "comma-separated glob patterns on relative paths to include (POSIX slash style)")
		excludeGlobCSV          = fset.String("exclude-glob", "", "comma-separated glob patterns on relative paths to exclude (POSIX slash style)")
		modeStr                 = fset.String("mode", "lenient", "parse mode: strict|lenient")
		reportFiles             = fset.String("report-files", "all", "report file entries policy: all|failed|issues|warnings|none")
		baselinePath            = fset.String("baseline", "", "path to baseline replay report JSON for regression diff")
		baselineLatest          = fset.Bool("baseline-latest", false, "use latest replay report under trend-dir as baseline")
		runID                   = fset.String("run-id", "", "optional run identifier (for trend output, e.g. git SHA)")
		trendDir                = fset.String("trend-dir", "", "directory containing historical replay report JSON files for trend summary")
		trendLimit              = fset.Int("trend-limit", 20, "max historical points to keep for trend, <=0 means unlimited")
		saveTrend               = fset.Bool("save-trend", false, "save current replay report JSON into trend-dir")
		saveTrendName           = fset.String("save-trend-name", "", "optional trend report file name when save-trend is enabled")
		saveTrendPrune          = fset.Bool("save-trend-prune", false, "when saving trend report, prune oldest trend JSON files to trend-limit")
		includeRaw              = fset.Bool("include-raw", false, "include raw artifact payloads in extraction")
		detectImages            = fset.Bool("detect-images", true, "enable image signature detection")
		detectOLEDS             = fset.Bool("detect-oleds", true, "enable OLEDS stream detection")
		unwrapOle10             = fset.Bool("unwrap-ole10native", true, "enable recursive Ole10Native unwrapping")
		dedup                   = fset.Bool("deduplicate", true, "enable SHA-256 dedup")
		maxDepth                = fset.Int("max-depth", 16, "max recursive extraction depth")
		maxArtifacts            = fset.Int("max-artifacts", 4096, "max artifacts per file")
		maxTotalBytes           = fset.Int64("max-total-bytes", 64<<20, "max total extracted bytes per file")
		maxArtifactSize         = fset.Int64("max-artifact-size", 32<<20, "max single artifact size in bytes")
		minProcessed            = fset.Int("min-processed", -1, "gate: minimum required processed files, negative disables")
		maxProcessed            = fset.Int("max-processed", -1, "gate: maximum allowed processed files, negative disables")
		minPassRate             = fset.Float64("min-pass-rate", -1, "gate: minimum acceptable pass rate in [0,1], negative disables")
		maxFailed               = fset.Int("max-failed", -1, "gate: maximum allowed failed files, negative disables")
		maxPartial              = fset.Int("max-partial", -1, "gate: maximum allowed partial files, negative disables")
		maxWarnings             = fset.Int("max-warnings", -1, "gate: maximum allowed warning count, negative disables")
		maxNewlyFailed          = fset.Int("max-newly-failed", -1, "gate: maximum allowed newly failed files vs baseline, negative disables")
		maxNewFiles             = fset.Int("max-new-files", -1, "gate: maximum allowed new files vs baseline, negative disables")
		maxRemovedFiles         = fset.Int("max-removed-files", -1, "gate: maximum allowed removed files vs baseline, negative disables")
		maxNewlyPartial         = fset.Int("max-newly-partial", -1, "gate: maximum allowed newly partial files vs baseline, negative disables")
		maxPassRateDrop         = fset.Float64("max-pass-rate-drop", -1, "gate: maximum allowed pass-rate drop vs latest trend point in [0,1], negative disables")
		maxFailedIncrease       = fset.Int("max-failed-increase", -1, "gate: maximum allowed failed-files increase vs latest trend point, negative disables")
		maxPartialIncrease      = fset.Int("max-partial-increase", -1, "gate: maximum allowed partial-files increase vs latest trend point, negative disables")
		maxWarningIncrease      = fset.Int("max-warning-increase", -1, "gate: maximum allowed warning-count increase vs latest trend point, negative disables")
		denyErrorCodes          = fset.String("deny-error-codes", "", "gate: comma-separated error codes that must not appear")
		maxNewErrorCodes        = fset.Int("max-new-error-codes", -1, "gate: maximum allowed new error codes vs baseline, negative disables")
		maxErrorCodeRegressions = fset.Int("max-error-code-regressions", -1, "gate: maximum allowed error codes with increased failure count vs baseline, negative disables")
		outputPath              = fset.String("output", "", "output report path; empty prints JSON to stdout")
	)
	if err := fset.Parse(args); err != nil {
		return err
	}
	baselinePathTrim := strings.TrimSpace(*baselinePath)
	trendDirTrim := strings.TrimSpace(*trendDir)
	if *baselineLatest && baselinePathTrim != "" {
		return errors.New("baseline-latest cannot be used with -baseline")
	}
	if *baselineLatest && trendDirTrim == "" {
		return errors.New("baseline-latest requires -trend-dir")
	}
	hasBaselineInput := baselinePathTrim != "" || *baselineLatest
	if *maxNewlyFailed >= 0 && !hasBaselineInput {
		return errors.New("max-newly-failed requires -baseline")
	}
	if *maxNewFiles >= 0 && !hasBaselineInput {
		return errors.New("max-new-files requires -baseline")
	}
	if *maxRemovedFiles >= 0 && !hasBaselineInput {
		return errors.New("max-removed-files requires -baseline")
	}
	if *maxNewlyPartial >= 0 && !hasBaselineInput {
		return errors.New("max-newly-partial requires -baseline")
	}
	if *maxNewErrorCodes >= 0 && !hasBaselineInput {
		return errors.New("max-new-error-codes requires -baseline")
	}
	if *maxErrorCodeRegressions >= 0 && !hasBaselineInput {
		return errors.New("max-error-code-regressions requires -baseline")
	}
	if *maxPassRateDrop >= 0 && trendDirTrim == "" {
		return errors.New("max-pass-rate-drop requires -trend-dir")
	}
	if *maxFailedIncrease >= 0 && trendDirTrim == "" {
		return errors.New("max-failed-increase requires -trend-dir")
	}
	if *maxPartialIncrease >= 0 && trendDirTrim == "" {
		return errors.New("max-partial-increase requires -trend-dir")
	}
	if *maxWarningIncrease >= 0 && trendDirTrim == "" {
		return errors.New("max-warning-increase requires -trend-dir")
	}
	if *saveTrend && trendDirTrim == "" {
		return errors.New("save-trend requires -trend-dir")
	}
	if *saveTrendPrune && !*saveTrend {
		return errors.New("save-trend-prune requires -save-trend")
	}
	if *saveTrendPrune && *trendLimit <= 0 {
		return errors.New("save-trend-prune requires trend-limit > 0")
	}
	if *minPassRate > 1 {
		return errors.New("min-pass-rate must be <= 1")
	}
	if *maxPassRateDrop > 1 {
		return errors.New("max-pass-rate-drop must be <= 1")
	}
	denyCodes := parseCSVTokens(*denyErrorCodes)
	reportFilesPolicy, err := parseReportFilesPolicy(*reportFiles)
	if err != nil {
		return err
	}
	includeGlobs, err := parsePathGlobs(*includeGlobCSV, "include-glob")
	if err != nil {
		return err
	}
	excludeGlobs, err := parsePathGlobs(*excludeGlobCSV, "exclude-glob")
	if err != nil {
		return err
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
		rel := path
		if r, relErr := filepath.Rel(absRoot, path); relErr == nil {
			rel = filepath.ToSlash(r)
		}
		if matchesExt(path, extensions) && matchesPathFilters(rel, includeGlobs, excludeGlobs) {
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
		IncludeGlobs:    append([]string(nil), includeGlobs...),
		ExcludeGlobs:    append([]string(nil), excludeGlobs...),
		Mode:            strings.ToLower(*modeStr),
		ReportFiles:     reportFilesPolicy,
		RunID:           strings.TrimSpace(*runID),
		BaselineReport:  baselinePathTrim,
		BaselineLatest:  *baselineLatest,
		TrendDir:        trendDirTrim,
		TrendLimit:      *trendLimit,
		SaveTrend:       *saveTrend,
		SaveTrendName:   strings.TrimSpace(*saveTrendName),
		SaveTrendPrune:  *saveTrendPrune,
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
			report.Summary.WarningsTotal += item.Warnings
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
	report.Summary.ErrorCodes = collectErrorCodeCounts(report.Files)
	resolvedBaselinePath := baselinePathTrim
	var baseRep replayReport
	baseLoaded := false
	if *baselineLatest {
		latestPath, latestRep, loadErr := loadLatestTrendReport(trendDirTrim)
		if loadErr != nil {
			return loadErr
		}
		resolvedBaselinePath = latestPath
		baseRep = latestRep
		baseLoaded = true
	} else if resolvedBaselinePath != "" {
		loaded, loadErr := loadReplayReport(resolvedBaselinePath)
		if loadErr != nil {
			return loadErr
		}
		baseRep = loaded
		baseLoaded = true
	}
	if baseLoaded {
		report.Baseline = diffReplayReport(resolvedBaselinePath, baseRep, report)
		report.Options.BaselineReport = resolvedBaselinePath
	}
	if trendDirTrim != "" {
		trend, trendErr := buildTrend(trendDirTrim, *trendLimit, report)
		if trendErr != nil {
			return trendErr
		}
		report.Trend = trend
	}

	var minPassRatePtr *float64
	if *minPassRate >= 0 {
		v := *minPassRate
		minPassRatePtr = &v
	}
	var minProcessedPtr *int
	if *minProcessed >= 0 {
		v := *minProcessed
		minProcessedPtr = &v
	}
	var maxProcessedPtr *int
	if *maxProcessed >= 0 {
		v := *maxProcessed
		maxProcessedPtr = &v
	}
	var maxFailedPtr *int
	if *maxFailed >= 0 {
		v := *maxFailed
		maxFailedPtr = &v
	}
	var maxPartialPtr *int
	if *maxPartial >= 0 {
		v := *maxPartial
		maxPartialPtr = &v
	}
	var maxWarningsPtr *int
	if *maxWarnings >= 0 {
		v := *maxWarnings
		maxWarningsPtr = &v
	}
	var maxNewlyFailedPtr *int
	if *maxNewlyFailed >= 0 {
		v := *maxNewlyFailed
		maxNewlyFailedPtr = &v
	}
	var maxNewFilesPtr *int
	if *maxNewFiles >= 0 {
		v := *maxNewFiles
		maxNewFilesPtr = &v
	}
	var maxRemovedFilesPtr *int
	if *maxRemovedFiles >= 0 {
		v := *maxRemovedFiles
		maxRemovedFilesPtr = &v
	}
	var maxNewlyPartialPtr *int
	if *maxNewlyPartial >= 0 {
		v := *maxNewlyPartial
		maxNewlyPartialPtr = &v
	}
	var maxPassRateDropPtr *float64
	if *maxPassRateDrop >= 0 {
		v := *maxPassRateDrop
		maxPassRateDropPtr = &v
	}
	var maxFailedIncreasePtr *int
	if *maxFailedIncrease >= 0 {
		v := *maxFailedIncrease
		maxFailedIncreasePtr = &v
	}
	var maxPartialIncreasePtr *int
	if *maxPartialIncrease >= 0 {
		v := *maxPartialIncrease
		maxPartialIncreasePtr = &v
	}
	var maxWarningIncreasePtr *int
	if *maxWarningIncrease >= 0 {
		v := *maxWarningIncrease
		maxWarningIncreasePtr = &v
	}
	var maxNewErrorCodesPtr *int
	if *maxNewErrorCodes >= 0 {
		v := *maxNewErrorCodes
		maxNewErrorCodesPtr = &v
	}
	var maxErrorCodeRegressionsPtr *int
	if *maxErrorCodeRegressions >= 0 {
		v := *maxErrorCodeRegressions
		maxErrorCodeRegressionsPtr = &v
	}
	gateErr := evaluateGates(
		&report,
		minProcessedPtr,
		maxProcessedPtr,
		minPassRatePtr,
		maxFailedPtr,
		maxPartialPtr,
		maxWarningsPtr,
		maxNewlyFailedPtr,
		maxNewFilesPtr,
		maxRemovedFilesPtr,
		maxNewlyPartialPtr,
		maxPassRateDropPtr,
		maxFailedIncreasePtr,
		maxPartialIncreasePtr,
		maxWarningIncreasePtr,
		denyCodes,
		maxNewErrorCodesPtr,
		maxErrorCodeRegressionsPtr,
	)
	applyReportFilePolicy(&report, reportFilesPolicy)

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
		if *saveTrend {
			savedPath, saveErr := saveTrendReport(trendDirTrim, strings.TrimSpace(*saveTrendName), strings.TrimSpace(*runID), buf)
			if saveErr != nil {
				return saveErr
			}
			if *saveTrendPrune {
				if _, pruneErr := pruneTrendReports(trendDirTrim, *trendLimit, savedPath); pruneErr != nil {
					return pruneErr
				}
			}
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
	if *saveTrend {
		savedPath, saveErr := saveTrendReport(trendDirTrim, strings.TrimSpace(*saveTrendName), strings.TrimSpace(*runID), buf)
		if saveErr != nil {
			return saveErr
		}
		if *saveTrendPrune {
			if _, pruneErr := pruneTrendReports(trendDirTrim, *trendLimit, savedPath); pruneErr != nil {
				return pruneErr
			}
		}
	}
	return gateErr
}

func loadLatestTrendReport(dir string) (string, replayReport, error) {
	absDir, err := filepath.Abs(strings.TrimSpace(dir))
	if err != nil {
		return "", replayReport{}, err
	}
	info, statErr := os.Stat(absDir)
	if statErr != nil {
		return "", replayReport{}, statErr
	}
	if !info.IsDir() {
		return "", replayReport{}, fmt.Errorf("trend-dir is not a directory: %s", absDir)
	}
	found := false
	latestPath := ""
	var latestReport replayReport
	var latestWhen time.Time
	err = filepath.WalkDir(absDir, func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		if strings.ToLower(filepath.Ext(path)) != ".json" {
			return nil
		}
		rep, loadErr := loadReplayReport(path)
		if loadErr != nil {
			return nil
		}
		when := parseReportTime(rep.GeneratedAt)
		if when.IsZero() {
			if fi, fiErr := os.Stat(path); fiErr == nil {
				when = fi.ModTime()
			}
		}
		if !found || when.After(latestWhen) || (when.Equal(latestWhen) && path > latestPath) {
			found = true
			latestWhen = when
			latestPath = path
			latestReport = rep
		}
		return nil
	})
	if err != nil {
		return "", replayReport{}, err
	}
	if !found {
		return "", replayReport{}, fmt.Errorf("no replay report found in trend-dir: %s", absDir)
	}
	return latestPath, latestReport, nil
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
		BaselinePath:        baselinePath,
		BaselineGenerated:   base.GeneratedAt,
		BaselineFiles:       len(base.Files),
		CurrentFiles:        len(cur.Files),
		ErrorCodeDelta:      map[string]int{},
		NewErrorCodes:       make([]string, 0, 8),
		IncreasedErrorCodes: make([]replayCodeDelta, 0, 8),
		Regressions:         make([]replayRegression, 0, 16),
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

	baseCodes := collectErrorCodeCounts(base.Files)
	curCodes := collectErrorCodeCounts(cur.Files)
	keys := sortedUnionKeys(baseCodes, curCodes)
	for _, k := range keys {
		delta := curCodes[k] - baseCodes[k]
		if delta != 0 {
			out.ErrorCodeDelta[k] = delta
		}
		if baseCodes[k] == 0 && curCodes[k] > 0 {
			out.NewErrorCodes = append(out.NewErrorCodes, k)
		}
		if delta > 0 {
			out.IncreasedErrorCodes = append(out.IncreasedErrorCodes, replayCodeDelta{
				Code:     k,
				Baseline: baseCodes[k],
				Current:  curCodes[k],
				Delta:    delta,
			})
		}
	}

	sort.Slice(out.Regressions, func(i, j int) bool {
		if out.Regressions[i].Path == out.Regressions[j].Path {
			return out.Regressions[i].Kind < out.Regressions[j].Kind
		}
		return out.Regressions[i].Path < out.Regressions[j].Path
	})
	sort.Strings(out.NewErrorCodes)
	sort.Slice(out.IncreasedErrorCodes, func(i, j int) bool {
		return out.IncreasedErrorCodes[i].Code < out.IncreasedErrorCodes[j].Code
	})
	if len(out.ErrorCodeDelta) == 0 {
		out.ErrorCodeDelta = nil
	}
	if len(out.NewErrorCodes) == 0 {
		out.NewErrorCodes = nil
	}
	if len(out.IncreasedErrorCodes) == 0 {
		out.IncreasedErrorCodes = nil
	}
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

func evaluateGates(
	report *replayReport,
	minProcessed *int,
	maxProcessed *int,
	minPassRate *float64,
	maxFailed *int,
	maxPartial *int,
	maxWarnings *int,
	maxNewlyFailed *int,
	maxNewFiles *int,
	maxRemovedFiles *int,
	maxNewlyPartial *int,
	maxPassRateDrop *float64,
	maxFailedIncrease *int,
	maxPartialIncrease *int,
	maxWarningIncrease *int,
	denyErrorCodes []string,
	maxNewErrorCodes *int,
	maxErrorCodeRegressions *int,
) error {
	if report == nil {
		return errors.New("nil report")
	}
	denyErrorCodes = normalizeCodes(denyErrorCodes)
	report.Gate = replayGateResult{
		Enabled:                 minProcessed != nil || maxProcessed != nil || minPassRate != nil || maxFailed != nil || maxPartial != nil || maxWarnings != nil || maxNewlyFailed != nil || maxNewFiles != nil || maxRemovedFiles != nil || maxNewlyPartial != nil || maxPassRateDrop != nil || maxFailedIncrease != nil || maxPartialIncrease != nil || maxWarningIncrease != nil || len(denyErrorCodes) > 0 || maxNewErrorCodes != nil || maxErrorCodeRegressions != nil,
		Passed:                  true,
		MinProcessed:            minProcessed,
		MaxProcessed:            maxProcessed,
		MinPassRate:             minPassRate,
		MaxFailed:               maxFailed,
		MaxPartial:              maxPartial,
		MaxWarnings:             maxWarnings,
		MaxNewlyFailed:          maxNewlyFailed,
		MaxNewFiles:             maxNewFiles,
		MaxRemovedFiles:         maxRemovedFiles,
		MaxNewlyPartial:         maxNewlyPartial,
		MaxPassRateDrop:         maxPassRateDrop,
		MaxFailedIncrease:       maxFailedIncrease,
		MaxPartialIncrease:      maxPartialIncrease,
		MaxWarningIncrease:      maxWarningIncrease,
		DenyErrorCodes:          denyErrorCodes,
		MaxNewErrorCodes:        maxNewErrorCodes,
		MaxErrorCodeRegressions: maxErrorCodeRegressions,
		Failures:                nil,
	}
	if !report.Gate.Enabled {
		return nil
	}
	if minProcessed != nil && report.Summary.Processed < *minProcessed {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("processed %d < min_processed %d", report.Summary.Processed, *minProcessed))
	}
	if maxProcessed != nil && report.Summary.Processed > *maxProcessed {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("processed %d > max_processed %d", report.Summary.Processed, *maxProcessed))
	}
	if minPassRate != nil && report.Summary.PassRate < *minPassRate {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("pass_rate %.6f < min_pass_rate %.6f", report.Summary.PassRate, *minPassRate))
	}
	if maxFailed != nil && report.Summary.Failed > *maxFailed {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("failed %d > max_failed %d", report.Summary.Failed, *maxFailed))
	}
	if maxPartial != nil && report.Summary.Partial > *maxPartial {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("partial %d > max_partial %d", report.Summary.Partial, *maxPartial))
	}
	if maxWarnings != nil && report.Summary.WarningsTotal > *maxWarnings {
		report.Gate.Failures = append(report.Gate.Failures,
			fmt.Sprintf("warnings %d > max_warnings %d", report.Summary.WarningsTotal, *maxWarnings))
	}
	if maxNewlyFailed != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_newly_failed gate requires baseline diff")
		} else if report.Baseline.NewlyFailed > *maxNewlyFailed {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("newly_failed %d > max_newly_failed %d", report.Baseline.NewlyFailed, *maxNewlyFailed))
		}
	}
	if maxNewFiles != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_new_files gate requires baseline diff")
		} else if report.Baseline.NewFiles > *maxNewFiles {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("new_files %d > max_new_files %d", report.Baseline.NewFiles, *maxNewFiles))
		}
	}
	if maxRemovedFiles != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_removed_files gate requires baseline diff")
		} else if report.Baseline.RemovedFiles > *maxRemovedFiles {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("removed_files %d > max_removed_files %d", report.Baseline.RemovedFiles, *maxRemovedFiles))
		}
	}
	if maxNewlyPartial != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_newly_partial gate requires baseline diff")
		} else if report.Baseline.NewlyPartial > *maxNewlyPartial {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("newly_partial %d > max_newly_partial %d", report.Baseline.NewlyPartial, *maxNewlyPartial))
		}
	}
	if maxPassRateDrop != nil {
		if report.Trend == nil || report.Trend.LatestDelta == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_pass_rate_drop gate requires trend latest delta")
		} else {
			drop := -report.Trend.LatestDelta.DeltaPassRate
			if drop > *maxPassRateDrop {
				report.Gate.Failures = append(report.Gate.Failures,
					fmt.Sprintf("pass_rate_drop %.6f > max_pass_rate_drop %.6f", drop, *maxPassRateDrop))
			}
		}
	}
	if maxFailedIncrease != nil {
		if report.Trend == nil || report.Trend.LatestDelta == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_failed_increase gate requires trend latest delta")
		} else if report.Trend.LatestDelta.DeltaFailed > *maxFailedIncrease {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("failed_increase %d > max_failed_increase %d", report.Trend.LatestDelta.DeltaFailed, *maxFailedIncrease))
		}
	}
	if maxPartialIncrease != nil {
		if report.Trend == nil || report.Trend.LatestDelta == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_partial_increase gate requires trend latest delta")
		} else if report.Trend.LatestDelta.DeltaPartial > *maxPartialIncrease {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("partial_increase %d > max_partial_increase %d", report.Trend.LatestDelta.DeltaPartial, *maxPartialIncrease))
		}
	}
	if maxWarningIncrease != nil {
		if report.Trend == nil || report.Trend.LatestDelta == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_warning_increase gate requires trend latest delta")
		} else if report.Trend.LatestDelta.DeltaWarnings > *maxWarningIncrease {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("warning_increase %d > max_warning_increase %d", report.Trend.LatestDelta.DeltaWarnings, *maxWarningIncrease))
		}
	}
	for _, code := range denyErrorCodes {
		if n := report.Summary.ErrorCodes[code]; n > 0 {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("error_code %s present %d time(s)", code, n))
		}
	}
	if maxNewErrorCodes != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_new_error_codes gate requires baseline diff")
		} else if len(report.Baseline.NewErrorCodes) > *maxNewErrorCodes {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("new_error_codes %d > max_new_error_codes %d", len(report.Baseline.NewErrorCodes), *maxNewErrorCodes))
		}
	}
	if maxErrorCodeRegressions != nil {
		if report.Baseline == nil {
			report.Gate.Failures = append(report.Gate.Failures, "max_error_code_regressions gate requires baseline diff")
		} else if len(report.Baseline.IncreasedErrorCodes) > *maxErrorCodeRegressions {
			report.Gate.Failures = append(report.Gate.Failures,
				fmt.Sprintf("error_code_regressions %d > max_error_code_regressions %d", len(report.Baseline.IncreasedErrorCodes), *maxErrorCodeRegressions))
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
	parts := parseCSVTokens(csv)
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.ToLower(p)
		if !strings.HasPrefix(p, ".") {
			p = "." + p
		}
		out = append(out, p)
	}
	return out
}

func parseCSVTokens(csv string) []string {
	csv = strings.TrimSpace(csv)
	if csv == "" {
		return nil
	}
	parts := strings.Split(csv, ",")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p == "" {
			continue
		}
		out = append(out, p)
	}
	return out
}

func parseReportFilesPolicy(v string) (string, error) {
	switch strings.ToLower(strings.TrimSpace(v)) {
	case "", "all":
		return "all", nil
	case "failed":
		return "failed", nil
	case "issues":
		return "issues", nil
	case "warnings":
		return "warnings", nil
	case "none":
		return "none", nil
	default:
		return "", fmt.Errorf("invalid report-files %q, expected all|failed|issues|warnings|none", v)
	}
}

func parsePathGlobs(csv, flagName string) ([]string, error) {
	patterns := parseCSVTokens(csv)
	if len(patterns) == 0 {
		return nil, nil
	}
	out := make([]string, 0, len(patterns))
	for _, p := range patterns {
		p = filepath.ToSlash(strings.TrimSpace(p))
		if p == "" {
			continue
		}
		if _, err := path.Match(p, "probe/path"); err != nil {
			return nil, fmt.Errorf("invalid %s pattern %q: %w", flagName, p, err)
		}
		out = append(out, p)
	}
	return out, nil
}

func matchesPathFilters(rel string, include, exclude []string) bool {
	rel = filepath.ToSlash(strings.TrimSpace(rel))
	if len(include) > 0 {
		matched := false
		for _, p := range include {
			ok, _ := path.Match(p, rel)
			if ok {
				matched = true
				break
			}
		}
		if !matched {
			return false
		}
	}
	for _, p := range exclude {
		ok, _ := path.Match(p, rel)
		if ok {
			return false
		}
	}
	return true
}

func applyReportFilePolicy(report *replayReport, policy string) {
	if report == nil {
		return
	}
	total := len(report.Files)
	switch policy {
	case "all":
		// keep all entries
	case "failed":
		filtered := make([]replayFileResult, 0, len(report.Files))
		for _, f := range report.Files {
			if !f.Success {
				filtered = append(filtered, f)
			}
		}
		report.Files = filtered
	case "issues":
		filtered := make([]replayFileResult, 0, len(report.Files))
		for _, f := range report.Files {
			if !f.Success || f.Partial || f.Warnings > 0 {
				filtered = append(filtered, f)
			}
		}
		report.Files = filtered
	case "warnings":
		filtered := make([]replayFileResult, 0, len(report.Files))
		for _, f := range report.Files {
			if f.Warnings > 0 {
				filtered = append(filtered, f)
			}
		}
		report.Files = filtered
	case "none":
		report.Files = nil
	default:
		// unreachable if parsed through parseReportFilesPolicy
	}
	report.Summary.ReportedFiles = len(report.Files)
	report.Summary.OmittedFiles = total - len(report.Files)
}

func normalizeCodes(codes []string) []string {
	if len(codes) == 0 {
		return nil
	}
	seen := map[string]struct{}{}
	out := make([]string, 0, len(codes))
	for _, c := range codes {
		c = strings.ToUpper(strings.TrimSpace(c))
		if c == "" {
			continue
		}
		if _, ok := seen[c]; ok {
			continue
		}
		seen[c] = struct{}{}
		out = append(out, c)
	}
	sort.Strings(out)
	return out
}

func collectErrorCodeCounts(files []replayFileResult) map[string]int {
	out := map[string]int{}
	for _, f := range files {
		if f.Success {
			continue
		}
		code := strings.TrimSpace(f.ErrorCode)
		if code == "" {
			code = "UNKNOWN"
		}
		code = strings.ToUpper(code)
		out[code]++
	}
	if len(out) == 0 {
		return nil
	}
	return out
}

func sortedUnionKeys(a, b map[string]int) []string {
	seen := map[string]struct{}{}
	out := make([]string, 0, len(a)+len(b))
	for k := range a {
		if _, ok := seen[k]; ok {
			continue
		}
		seen[k] = struct{}{}
		out = append(out, k)
	}
	for k := range b {
		if _, ok := seen[k]; ok {
			continue
		}
		seen[k] = struct{}{}
		out = append(out, k)
	}
	sort.Strings(out)
	return out
}

func buildTrend(dir string, limit int, current replayReport) (*replayTrend, error) {
	absDir, err := filepath.Abs(dir)
	if err != nil {
		return nil, err
	}
	if info, statErr := os.Stat(absDir); statErr != nil {
		if errors.Is(statErr, os.ErrNotExist) {
			point := reportToTrendPoint(current, "current")
			return &replayTrend{
				SourceDir:   absDir,
				Points:      []replayTrendPoint{point},
				LatestDelta: nil,
			}, nil
		}
		return nil, statErr
	} else if !info.IsDir() {
		return nil, fmt.Errorf("trend-dir is not a directory: %s", absDir)
	}
	entries := make([]trendEntry, 0, 32)
	err = filepath.WalkDir(absDir, func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		if strings.ToLower(filepath.Ext(path)) != ".json" {
			return nil
		}
		rep, loadErr := loadReplayReport(path)
		if loadErr != nil {
			return nil
		}
		tm := parseReportTime(rep.GeneratedAt)
		if tm.IsZero() {
			if info, statErr := os.Stat(path); statErr == nil {
				tm = info.ModTime()
			}
		}
		entries = append(entries, trendEntry{
			Path:   path,
			Report: rep,
			When:   tm,
		})
		return nil
	})
	if err != nil {
		return nil, err
	}
	sort.Slice(entries, func(i, j int) bool {
		if entries[i].When.Equal(entries[j].When) {
			return entries[i].Path < entries[j].Path
		}
		return entries[i].When.Before(entries[j].When)
	})
	if limit > 0 && len(entries) > limit {
		entries = entries[len(entries)-limit:]
	}

	points := make([]replayTrendPoint, 0, len(entries)+1)
	for _, e := range entries {
		rel := e.Path
		if r, relErr := filepath.Rel(absDir, e.Path); relErr == nil {
			rel = filepath.ToSlash(r)
		}
		points = append(points, reportToTrendPoint(e.Report, rel))
	}
	points = append(points, reportToTrendPoint(current, "current"))

	out := &replayTrend{
		SourceDir: absDir,
		Points:    points,
	}
	if len(points) >= 2 {
		prev := points[len(points)-2]
		last := points[len(points)-1]
		out.LatestDelta = &replayTrendDelta{
			FromGeneratedAt: prev.GeneratedAt,
			ToGeneratedAt:   last.GeneratedAt,
			DeltaPassRate:   last.PassRate - prev.PassRate,
			DeltaFailed:     last.Failed - prev.Failed,
			DeltaPartial:    last.Partial - prev.Partial,
			DeltaWarnings:   last.Warnings - prev.Warnings,
		}
	}
	return out, nil
}

type trendEntry struct {
	Path   string
	Report replayReport
	When   time.Time
}

func parseReportTime(v string) time.Time {
	v = strings.TrimSpace(v)
	if v == "" {
		return time.Time{}
	}
	tm, err := time.Parse(time.RFC3339Nano, v)
	if err != nil {
		return time.Time{}
	}
	return tm
}

func reportToTrendPoint(rep replayReport, reportPath string) replayTrendPoint {
	return replayTrendPoint{
		GeneratedAt: rep.GeneratedAt,
		RunID:       rep.Options.RunID,
		ReportPath:  reportPath,
		Processed:   rep.Summary.Processed,
		Success:     rep.Summary.Success,
		Failed:      rep.Summary.Failed,
		Partial:     rep.Summary.Partial,
		Warnings:    rep.Summary.WarningsTotal,
		PassRate:    rep.Summary.PassRate,
	}
}

func saveTrendReport(dir, name, runID string, buf []byte) (string, error) {
	absDir, err := filepath.Abs(strings.TrimSpace(dir))
	if err != nil {
		return "", err
	}
	if err := os.MkdirAll(absDir, 0o755); err != nil {
		return "", err
	}
	fileName := strings.TrimSpace(name)
	if fileName == "" {
		fileName = buildTrendFileName(runID)
	}
	fileName = filepath.Base(fileName)
	if strings.TrimSpace(fileName) == "" || fileName == "." || fileName == ".." {
		return "", errors.New("invalid save-trend-name")
	}
	if !strings.HasSuffix(strings.ToLower(fileName), ".json") {
		fileName += ".json"
	}
	path := filepath.Join(absDir, fileName)
	if err := os.WriteFile(path, buf, 0o644); err != nil {
		return "", err
	}
	return path, nil
}

func buildTrendFileName(runID string) string {
	stamp := time.Now().UTC().Format("20060102T150405Z")
	run := sanitizeFileToken(runID)
	if run == "" {
		return "replay-" + stamp + ".json"
	}
	return "replay-" + stamp + "-" + run + ".json"
}

func pruneTrendReports(dir string, keep int, preservePath string) (int, error) {
	if keep <= 0 {
		return 0, nil
	}
	absDir, err := filepath.Abs(strings.TrimSpace(dir))
	if err != nil {
		return 0, err
	}
	preserveAbs := ""
	if strings.TrimSpace(preservePath) != "" {
		preserveAbs, _ = filepath.Abs(strings.TrimSpace(preservePath))
	}
	entries := make([]trendEntry, 0, keep+8)
	err = filepath.WalkDir(absDir, func(path string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if d.IsDir() {
			return nil
		}
		if strings.ToLower(filepath.Ext(path)) != ".json" {
			return nil
		}
		rep, loadErr := loadReplayReport(path)
		if loadErr != nil {
			return nil
		}
		when := parseReportTime(rep.GeneratedAt)
		if when.IsZero() {
			if info, statErr := os.Stat(path); statErr == nil {
				when = info.ModTime()
			}
		}
		entries = append(entries, trendEntry{
			Path:   path,
			Report: rep,
			When:   when,
		})
		return nil
	})
	if err != nil {
		return 0, err
	}
	if len(entries) <= keep {
		return 0, nil
	}
	sort.Slice(entries, func(i, j int) bool {
		if entries[i].When.Equal(entries[j].When) {
			return entries[i].Path < entries[j].Path
		}
		return entries[i].When.Before(entries[j].When)
	})

	toDelete := len(entries) - keep
	removed := 0
	for i := 0; i < len(entries) && toDelete > 0; i++ {
		p := entries[i].Path
		pAbs, _ := filepath.Abs(p)
		if preserveAbs != "" && strings.EqualFold(pAbs, preserveAbs) {
			continue
		}
		if err := os.Remove(p); err != nil && !errors.Is(err, os.ErrNotExist) {
			return removed, err
		}
		removed++
		toDelete--
	}
	return removed, nil
}

func sanitizeFileToken(v string) string {
	v = strings.TrimSpace(v)
	if v == "" {
		return ""
	}
	var b strings.Builder
	for _, r := range v {
		if (r >= 'a' && r <= 'z') || (r >= 'A' && r <= 'Z') || (r >= '0' && r <= '9') || r == '_' || r == '-' || r == '.' {
			b.WriteRune(r)
			continue
		}
		b.WriteByte('_')
	}
	out := strings.Trim(b.String(), "._- ")
	if len(out) > 64 {
		out = out[:64]
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
