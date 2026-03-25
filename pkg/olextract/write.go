package olextract

import (
	"crypto/rand"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"strconv"
	"strings"
	"time"
	"unicode"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
	"github.com/KairuiKin/go-olespec/pkg/olecfb/storage"
)

type WriteOptions struct {
	Overwrite         bool
	Layout            WriteLayout
	WriteManifest     bool
	ManifestName      string
	AtomicPublish     bool
	PreferOLEFileName bool
}

type WrittenFile struct {
	ArtifactID   string
	ArtifactPath string
	RelativePath string
	FilePath     string
	Size         int64
}

type WriteResult struct {
	FilesWritten int
	BytesWritten int64
	Skipped      int
	ManifestPath string
	Files        []WrittenFile
}

type WriteLayout string

const (
	WriteLayoutFlat WriteLayout = "flat"
	WriteLayoutTree WriteLayout = "tree"
)

type writePlan struct {
	artifact olecfb.Artifact
	rel      string
	abs      string
}

type committedRecord struct {
	file       WrittenFile
	backupPath string
}

var writeFile = os.WriteFile
var renameFile = os.Rename

// WriteArtifacts writes artifact raw payloads to a directory with deterministic paths.
// Artifacts without Raw payload are skipped.
func WriteArtifacts(report *olecfb.ExtractReport, dstDir string, opt WriteOptions) (WriteResult, error) {
	if report == nil {
		return WriteResult{}, &olecfb.OLEError{
			Code:    olecfb.ErrInvalidArgument,
			Message: "extract report is nil",
			Op:      "olextract.write_artifacts",
			Offset:  -1,
		}
	}
	if strings.TrimSpace(dstDir) == "" {
		return WriteResult{}, &olecfb.OLEError{
			Code:    olecfb.ErrInvalidArgument,
			Message: "destination directory is empty",
			Op:      "olextract.write_artifacts",
			Offset:  -1,
		}
	}
	if err := os.MkdirAll(dstDir, 0o755); err != nil {
		return WriteResult{}, &olecfb.OLEError{
			Code:    olecfb.ErrUnsupported,
			Message: "create destination directory failed",
			Op:      "olextract.write_artifacts",
			Offset:  -1,
			Cause:   err,
		}
	}

	layout := opt.Layout
	if layout == "" {
		layout = WriteLayoutFlat
	}
	if layout != WriteLayoutFlat && layout != WriteLayoutTree {
		return WriteResult{}, &olecfb.OLEError{
			Code:    olecfb.ErrInvalidArgument,
			Message: "unsupported write layout",
			Op:      "olextract.write_artifacts",
			Offset:  -1,
		}
	}
	out := WriteResult{}
	plans := make([]writePlan, 0, len(report.Artifacts))
	for i, a := range report.Artifacts {
		if len(a.Raw) == 0 {
			out.Skipped++
			continue
		}
		var rel string
		switch layout {
		case WriteLayoutTree:
			rel = buildArtifactTreePath(i, a, opt.PreferOLEFileName)
		default:
			rel = buildArtifactFileName(i, a, opt.PreferOLEFileName)
		}
		p := filepath.Join(dstDir, rel)
		plans = append(plans, writePlan{artifact: a, rel: rel, abs: p})
	}

	var manifestPath string
	if opt.WriteManifest {
		name, err := validateManifestName(opt.ManifestName)
		if err != nil {
			return WriteResult{}, err
		}
		manifestPath = filepath.Join(dstDir, name)
	}

	// Preflight conflicts before any writes to avoid partial output on conflict.
	if !opt.Overwrite {
		for _, plan := range plans {
			if _, err := os.Stat(plan.abs); err == nil {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrConflict,
					Message: "destination file already exists",
					Op:      "olextract.write_artifacts",
					Path:    plan.abs,
					Offset:  -1,
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrUnsupported,
					Message: "check destination file failed",
					Op:      "olextract.write_artifacts",
					Path:    plan.abs,
					Offset:  -1,
					Cause:   err,
				}
			}
		}
		if manifestPath != "" {
			if _, err := os.Stat(manifestPath); err == nil {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrConflict,
					Message: "manifest file already exists",
					Op:      "olextract.write_artifacts",
					Path:    manifestPath,
					Offset:  -1,
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrUnsupported,
					Message: "check manifest file failed",
					Op:      "olextract.write_artifacts",
					Path:    manifestPath,
					Offset:  -1,
					Cause:   err,
				}
			}
		}
	}

	if opt.AtomicPublish {
		var err error
		out, err = writeArtifactsAtomic(dstDir, plans, report, manifestPath, out.Skipped, opt.Overwrite)
		if err != nil {
			return WriteResult{}, err
		}
		return out, nil
	}

	for _, plan := range plans {
		if err := os.MkdirAll(filepath.Dir(plan.abs), 0o755); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "create artifact sub directory failed",
				Op:      "olextract.write_artifacts",
				Path:    filepath.Dir(plan.abs),
				Offset:  -1,
				Cause:   err,
			}
		}
		if err := writeFile(plan.abs, plan.artifact.Raw, 0o644); err != nil {
			if !opt.Overwrite {
				rollbackWrittenFiles(out.Files)
			}
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "write artifact file failed",
				Op:      "olextract.write_artifacts",
				Path:    plan.abs,
				Offset:  -1,
				Cause:   err,
			}
		}
		out.FilesWritten++
		out.BytesWritten += int64(len(plan.artifact.Raw))
		out.Files = append(out.Files, WrittenFile{
			ArtifactID:   plan.artifact.ID,
			ArtifactPath: plan.artifact.Path,
			RelativePath: plan.rel,
			FilePath:     plan.abs,
			Size:         int64(len(plan.artifact.Raw)),
		})
	}
	if manifestPath != "" {
		if err := writeManifestFile(manifestPath, report, out); err != nil {
			if !opt.Overwrite {
				rollbackWrittenFiles(out.Files)
				_ = os.Remove(manifestPath)
			}
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "write manifest failed",
				Op:      "olextract.write_artifacts",
				Path:    manifestPath,
				Offset:  -1,
				Cause:   err,
			}
		}
		out.ManifestPath = manifestPath
	}
	return out, nil
}

func writeArtifactsAtomic(dstDir string, plans []writePlan, report *olecfb.ExtractReport, manifestPath string, skipped int, overwrite bool) (WriteResult, error) {
	stageDir := filepath.Join(dstDir, buildStageDirName())
	if err := os.MkdirAll(stageDir, 0o755); err != nil {
		return WriteResult{}, &olecfb.OLEError{
			Code:    olecfb.ErrUnsupported,
			Message: "create stage directory failed",
			Op:      "olextract.write_artifacts",
			Path:    stageDir,
			Offset:  -1,
			Cause:   err,
		}
	}
	defer os.RemoveAll(stageDir)

	out := WriteResult{Skipped: skipped}
	stageAbsByRel := map[string]string{}
	for _, plan := range plans {
		stageAbs := filepath.Join(stageDir, plan.rel)
		stageAbsByRel[plan.rel] = stageAbs
		if err := os.MkdirAll(filepath.Dir(stageAbs), 0o755); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "create stage sub directory failed",
				Op:      "olextract.write_artifacts",
				Path:    filepath.Dir(stageAbs),
				Offset:  -1,
				Cause:   err,
			}
		}
		if err := writeFile(stageAbs, plan.artifact.Raw, 0o644); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "write stage artifact file failed",
				Op:      "olextract.write_artifacts",
				Path:    stageAbs,
				Offset:  -1,
				Cause:   err,
			}
		}
	}

	stageManifest := ""
	if manifestPath != "" {
		stageManifest = filepath.Join(stageDir, filepath.Base(manifestPath))
		preview := WriteResult{
			FilesWritten: len(plans),
			BytesWritten: 0,
			Skipped:      skipped,
			Files:        make([]WrittenFile, 0, len(plans)),
		}
		for _, plan := range plans {
			sz := int64(len(plan.artifact.Raw))
			preview.BytesWritten += sz
			preview.Files = append(preview.Files, WrittenFile{
				ArtifactID:   plan.artifact.ID,
				ArtifactPath: plan.artifact.Path,
				RelativePath: plan.rel,
				FilePath:     plan.abs,
				Size:         sz,
			})
		}
		if err := writeManifestFile(stageManifest, report, preview); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "write stage manifest failed",
				Op:      "olextract.write_artifacts",
				Path:    stageManifest,
				Offset:  -1,
				Cause:   err,
			}
		}
	}

	committed := make([]committedRecord, 0, len(plans))
	for _, plan := range plans {
		stageAbs := stageAbsByRel[plan.rel]
		if err := os.MkdirAll(filepath.Dir(plan.abs), 0o755); err != nil {
			rollbackCommitted(committed)
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrCommitFailed,
				Message: "create destination sub directory failed during atomic publish",
				Op:      "olextract.write_artifacts",
				Path:    filepath.Dir(plan.abs),
				Offset:  -1,
				Cause:   err,
			}
		}
		backupPath := ""
		if overwrite {
			if _, err := os.Stat(plan.abs); err == nil {
				backupPath = filepath.Join(stageDir, ".bak-"+strconv.Itoa(len(committed)))
				if err := renameFile(plan.abs, backupPath); err != nil {
					rollbackCommitted(committed)
					return WriteResult{}, &olecfb.OLEError{
						Code:    olecfb.ErrCommitFailed,
						Message: "atomic publish backup rename failed",
						Op:      "olextract.write_artifacts",
						Path:    plan.abs,
						Offset:  -1,
						Cause:   err,
					}
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				rollbackCommitted(committed)
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrCommitFailed,
					Message: "atomic publish destination stat failed",
					Op:      "olextract.write_artifacts",
					Path:    plan.abs,
					Offset:  -1,
					Cause:   err,
				}
			}
		}
		if err := renameFile(stageAbs, plan.abs); err != nil {
			if backupPath != "" {
				_ = renameFile(backupPath, plan.abs)
			}
			rollbackCommitted(committed)
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrCommitFailed,
				Message: "atomic publish rename failed",
				Op:      "olextract.write_artifacts",
				Path:    plan.abs,
				Offset:  -1,
				Cause:   err,
			}
		}
		wf := WrittenFile{
			ArtifactID:   plan.artifact.ID,
			ArtifactPath: plan.artifact.Path,
			RelativePath: plan.rel,
			FilePath:     plan.abs,
			Size:         int64(len(plan.artifact.Raw)),
		}
		committed = append(committed, committedRecord{
			file:       wf,
			backupPath: backupPath,
		})
		out.Files = append(out.Files, wf)
		out.FilesWritten++
		out.BytesWritten += wf.Size
	}

	if manifestPath != "" {
		manifestBackup := ""
		if overwrite {
			if _, err := os.Stat(manifestPath); err == nil {
				manifestBackup = filepath.Join(stageDir, ".bak-manifest")
				if err := renameFile(manifestPath, manifestBackup); err != nil {
					rollbackCommitted(committed)
					return WriteResult{}, &olecfb.OLEError{
						Code:    olecfb.ErrCommitFailed,
						Message: "atomic publish manifest backup rename failed",
						Op:      "olextract.write_artifacts",
						Path:    manifestPath,
						Offset:  -1,
						Cause:   err,
					}
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				rollbackCommitted(committed)
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrCommitFailed,
					Message: "atomic publish manifest stat failed",
					Op:      "olextract.write_artifacts",
					Path:    manifestPath,
					Offset:  -1,
					Cause:   err,
				}
			}
		}
		if err := renameFile(stageManifest, manifestPath); err != nil {
			if manifestBackup != "" {
				_ = renameFile(manifestBackup, manifestPath)
			}
			rollbackCommitted(committed)
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrCommitFailed,
				Message: "atomic publish manifest rename failed",
				Op:      "olextract.write_artifacts",
				Path:    manifestPath,
				Offset:  -1,
				Cause:   err,
			}
		}
		out.ManifestPath = manifestPath
	}
	return out, nil
}

func buildStageDirName() string {
	return ".olespec-stage-" + strconv.FormatInt(time.Now().UnixNano(), 10) + "-" + randomHex4()
}

func randomHex4() string {
	var b [2]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "0000"
	}
	const hex = "0123456789abcdef"
	out := make([]byte, 4)
	out[0] = hex[b[0]>>4]
	out[1] = hex[b[0]&0x0F]
	out[2] = hex[b[1]>>4]
	out[3] = hex[b[1]&0x0F]
	return string(out)
}

func rollbackCommitted(records []committedRecord) {
	for i := len(records) - 1; i >= 0; i-- {
		_ = os.Remove(records[i].file.FilePath)
		if records[i].backupPath != "" {
			_ = renameFile(records[i].backupPath, records[i].file.FilePath)
		}
	}
}

func rollbackWrittenFiles(files []WrittenFile) {
	for i := len(files) - 1; i >= 0; i-- {
		_ = os.Remove(files[i].FilePath)
	}
}

func validateManifestName(name string) (string, error) {
	name = strings.TrimSpace(name)
	if name == "" {
		return "manifest.json", nil
	}
	if filepath.IsAbs(name) ||
		strings.Contains(name, "/") ||
		strings.Contains(name, "\\") ||
		name == "." ||
		name == ".." ||
		filepath.Base(name) != name {
		return "", &olecfb.OLEError{
			Code:    olecfb.ErrInvalidArgument,
			Message: "manifest name must be a file name without path separators",
			Op:      "olextract.write_artifacts",
			Offset:  -1,
		}
	}
	return name, nil
}

// ExtractFileToDir extracts artifacts from a file and writes raw payloads to dstDir.
// It always enables IncludeRaw for extraction.
func ExtractFileToDir(path, dstDir string, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions, writeOpt WriteOptions) (*olecfb.ExtractReport, WriteResult, error) {
	extractOpt.IncludeRaw = true
	rep, err := ExtractFile(path, openOpt, extractOpt)
	if err != nil {
		return nil, WriteResult{}, err
	}
	res, err := WriteArtifacts(rep, dstDir, writeOpt)
	if err != nil {
		return rep, WriteResult{}, err
	}
	return rep, res, nil
}

// ExtractBytesToDir extracts artifacts from a buffer and writes raw payloads to dstDir.
// It always enables IncludeRaw for extraction.
func ExtractBytesToDir(buf []byte, dstDir string, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions, writeOpt WriteOptions) (*olecfb.ExtractReport, WriteResult, error) {
	extractOpt.IncludeRaw = true
	rep, err := ExtractBytes(buf, openOpt, extractOpt)
	if err != nil {
		return nil, WriteResult{}, err
	}
	res, err := WriteArtifacts(rep, dstDir, writeOpt)
	if err != nil {
		return rep, WriteResult{}, err
	}
	return rep, res, nil
}

// ExtractReaderToDir extracts artifacts from a reader and writes raw payloads to dstDir.
// It always enables IncludeRaw for extraction.
func ExtractReaderToDir(r io.Reader, dstDir string, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions, writeOpt WriteOptions) (*olecfb.ExtractReport, WriteResult, error) {
	extractOpt.IncludeRaw = true
	rep, err := ExtractReader(r, openOpt, extractOpt)
	if err != nil {
		return nil, WriteResult{}, err
	}
	res, err := WriteArtifacts(rep, dstDir, writeOpt)
	if err != nil {
		return rep, WriteResult{}, err
	}
	return rep, res, nil
}

// ExtractBackendToDir extracts artifacts from a backend and writes raw payloads to dstDir.
// It always enables IncludeRaw for extraction.
func ExtractBackendToDir(rb storage.ReadBackend, dstDir string, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions, writeOpt WriteOptions) (*olecfb.ExtractReport, WriteResult, error) {
	extractOpt.IncludeRaw = true
	rep, err := ExtractBackend(rb, openOpt, extractOpt)
	if err != nil {
		return nil, WriteResult{}, err
	}
	res, err := WriteArtifacts(rep, dstDir, writeOpt)
	if err != nil {
		return rep, WriteResult{}, err
	}
	return rep, res, nil
}

func buildArtifactFileName(index int, a olecfb.Artifact, preferOLEFileName bool) string {
	base := preferredStem(a, preferOLEFileName)
	if base == "" {
		base = sanitizeName(a.Path)
	}
	if base == "" {
		base = "artifact"
	}
	if len(base) > 80 {
		base = base[:80]
	}
	ext := extensionByArtifact(a)
	return fmt.Sprintf("%06d_%s%s", index, base, ext)
}

func buildArtifactTreePath(index int, a olecfb.Artifact, preferOLEFileName bool) string {
	segments := splitArtifactPathForTree(a.Path)
	if len(segments) == 0 {
		segments = []string{"artifact"}
	}
	for i := range segments {
		segments[i] = sanitizeName(segments[i])
		if segments[i] == "" {
			segments[i] = "artifact"
		}
	}
	fileName := preferredStem(a, preferOLEFileName)
	if fileName == "" {
		fileName = segments[len(segments)-1]
	}
	if len(fileName) > 80 {
		fileName = fileName[:80]
	}
	fileName = fmt.Sprintf("%06d_%s%s", index, fileName, extensionByArtifact(a))
	if len(segments) == 1 {
		return fileName
	}
	dir := filepath.Join(segments[:len(segments)-1]...)
	return filepath.Join(dir, fileName)
}

func preferredStem(a olecfb.Artifact, enabled bool) string {
	if !enabled || a.OLEFileName == "" {
		return ""
	}
	base := path.Base(a.OLEFileName)
	ext := path.Ext(base)
	stem := strings.TrimSuffix(base, ext)
	return sanitizeName(stem)
}

func splitArtifactPathForTree(path string) []string {
	if strings.TrimSpace(path) == "" {
		return nil
	}
	path = strings.TrimPrefix(path, "/")
	path = strings.ReplaceAll(path, "!", "/_ole_/")
	raw := strings.Split(path, "/")
	out := make([]string, 0, len(raw))
	for _, s := range raw {
		if s == "" {
			continue
		}
		out = append(out, s)
	}
	return out
}

func sanitizeName(v string) string {
	var b strings.Builder
	for _, r := range v {
		if unicode.IsLetter(r) || unicode.IsDigit(r) || r == '.' || r == '_' || r == '-' {
			b.WriteRune(r)
			continue
		}
		b.WriteByte('_')
	}
	s := strings.Trim(b.String(), "_")
	s = strings.Trim(s, ". ")
	if s == "" {
		return s
	}
	if isWindowsReservedName(s) {
		return "_" + s
	}
	return s
}

func isWindowsReservedName(v string) bool {
	up := strings.ToUpper(v)
	base := up
	if i := strings.IndexByte(base, '.'); i >= 0 {
		base = base[:i]
	}
	switch base {
	case "CON", "PRN", "AUX", "NUL",
		"COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
		"LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9":
		return true
	default:
		return false
	}
}

type writeManifestData struct {
	FilesWritten int                    `json:"files_written"`
	BytesWritten int64                  `json:"bytes_written"`
	Skipped      int                    `json:"skipped"`
	Files        []writeManifestEntry   `json:"files"`
	Artifacts    []writeArtifactSummary `json:"artifacts"`
}

type writeManifestEntry struct {
	ArtifactID   string `json:"artifact_id"`
	ArtifactPath string `json:"artifact_path"`
	RelativePath string `json:"relative_path"`
	FilePath     string `json:"file_path"`
	Size         int64  `json:"size"`
}

type writeArtifactSummary struct {
	ID            string                `json:"id"`
	Path          string                `json:"path"`
	Kind          olecfb.ArtifactKind   `json:"kind"`
	Status        olecfb.ArtifactStatus `json:"status"`
	Size          int64                 `json:"size"`
	SHA256        string                `json:"sha256"`
	HasRaw        bool                  `json:"has_raw"`
	OLEFileName   string                `json:"ole_file_name,omitempty"`
	OLESourcePath string                `json:"ole_source_path,omitempty"`
	OLETempPath   string                `json:"ole_temp_path,omitempty"`
}

func writeManifestFile(path string, report *olecfb.ExtractReport, res WriteResult) error {
	m := writeManifestData{
		FilesWritten: res.FilesWritten,
		BytesWritten: res.BytesWritten,
		Skipped:      res.Skipped,
		Files:        make([]writeManifestEntry, 0, len(res.Files)),
		Artifacts:    make([]writeArtifactSummary, 0, len(report.Artifacts)),
	}
	for _, f := range res.Files {
		m.Files = append(m.Files, writeManifestEntry{
			ArtifactID:   f.ArtifactID,
			ArtifactPath: f.ArtifactPath,
			RelativePath: filepath.ToSlash(f.RelativePath),
			FilePath:     f.FilePath,
			Size:         f.Size,
		})
	}
	for _, a := range report.Artifacts {
		m.Artifacts = append(m.Artifacts, writeArtifactSummary{
			ID:            a.ID,
			Path:          a.Path,
			Kind:          a.Kind,
			Status:        a.Status,
			Size:          a.Size,
			SHA256:        a.SHA256,
			HasRaw:        len(a.Raw) > 0,
			OLEFileName:   a.OLEFileName,
			OLESourcePath: a.OLESourcePath,
			OLETempPath:   a.OLETempPath,
		})
	}
	buf, err := json.MarshalIndent(m, "", "  ")
	if err != nil {
		return err
	}
	return writeFile(path, buf, 0o644)
}

func extensionByArtifact(a olecfb.Artifact) string {
	if (a.Kind == olecfb.ArtifactOleObj || a.Kind == olecfb.ArtifactStream) && a.OLEFileName != "" {
		if ext, ok := safeExt(path.Ext(a.OLEFileName)); ok {
			return ext
		}
	}
	switch a.MediaType {
	case "image/png":
		return ".png"
	case "image/jpeg":
		return ".jpg"
	case "image/gif":
		return ".gif"
	case "image/bmp":
		return ".bmp"
	case "image/tiff":
		return ".tiff"
	case "image/webp":
		return ".webp"
	}
	switch a.Kind {
	case olecfb.ArtifactOLEFile:
		return ".ole"
	case olecfb.ArtifactOleObj:
		return ".oleobj"
	case olecfb.ArtifactImage:
		return ".img"
	default:
		return ".bin"
	}
}

func safeExt(ext string) (string, bool) {
	if ext == "" || len(ext) > 16 {
		return "", false
	}
	if ext[0] != '.' {
		return "", false
	}
	for _, r := range ext[1:] {
		if !(unicode.IsLetter(r) || unicode.IsDigit(r)) {
			return "", false
		}
	}
	return strings.ToLower(ext), true
}
