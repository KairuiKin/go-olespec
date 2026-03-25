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
		includeRaw      = fset.Bool("include-raw", false, "include raw artifact payloads in extraction")
		detectImages    = fset.Bool("detect-images", true, "enable image signature detection")
		detectOLEDS     = fset.Bool("detect-oleds", true, "enable OLEDS stream detection")
		unwrapOle10     = fset.Bool("unwrap-ole10native", true, "enable recursive Ole10Native unwrapping")
		dedup           = fset.Bool("deduplicate", true, "enable SHA-256 dedup")
		maxDepth        = fset.Int("max-depth", 16, "max recursive extraction depth")
		maxArtifacts    = fset.Int("max-artifacts", 4096, "max artifacts per file")
		maxTotalBytes   = fset.Int64("max-total-bytes", 64<<20, "max total extracted bytes per file")
		maxArtifactSize = fset.Int64("max-artifact-size", 32<<20, "max single artifact size in bytes")
		outputPath      = fset.String("output", "", "output report path; empty prints JSON to stdout")
	)
	if err := fset.Parse(args); err != nil {
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
		return err
	}

	outAbs, err := filepath.Abs(*outputPath)
	if err != nil {
		return err
	}
	if err := os.MkdirAll(filepath.Dir(outAbs), 0o755); err != nil {
		return err
	}
	return os.WriteFile(outAbs, buf, 0o644)
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
