package olextract

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
	"github.com/KairuiKin/go-olespec/pkg/olecfb/storage"
)

type WriteOptions struct {
	Overwrite     bool
	Layout        WriteLayout
	WriteManifest bool
	ManifestName  string
}

type WrittenFile struct {
	ArtifactID   string
	ArtifactPath string
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

// WriteArtifacts writes artifact raw payloads to a directory using deterministic flat file names.
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
	for i, a := range report.Artifacts {
		if len(a.Raw) == 0 {
			out.Skipped++
			continue
		}
		var rel string
		switch layout {
		case WriteLayoutTree:
			rel = buildArtifactTreePath(i, a)
		default:
			rel = buildArtifactFileName(i, a)
		}
		p := filepath.Join(dstDir, rel)
		if !opt.Overwrite {
			if _, err := os.Stat(p); err == nil {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrConflict,
					Message: "destination file already exists",
					Op:      "olextract.write_artifacts",
					Path:    p,
					Offset:  -1,
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrUnsupported,
					Message: "check destination file failed",
					Op:      "olextract.write_artifacts",
					Path:    p,
					Offset:  -1,
					Cause:   err,
				}
			}
		}
		if err := os.MkdirAll(filepath.Dir(p), 0o755); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "create artifact sub directory failed",
				Op:      "olextract.write_artifacts",
				Path:    filepath.Dir(p),
				Offset:  -1,
				Cause:   err,
			}
		}
		if err := os.WriteFile(p, a.Raw, 0o644); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "write artifact file failed",
				Op:      "olextract.write_artifacts",
				Path:    p,
				Offset:  -1,
				Cause:   err,
			}
		}
		out.FilesWritten++
		out.BytesWritten += int64(len(a.Raw))
		out.Files = append(out.Files, WrittenFile{
			ArtifactID:   a.ID,
			ArtifactPath: a.Path,
			FilePath:     p,
			Size:         int64(len(a.Raw)),
		})
	}
	if opt.WriteManifest {
		name := opt.ManifestName
		if strings.TrimSpace(name) == "" {
			name = "manifest.json"
		}
		mp := filepath.Join(dstDir, name)
		if !opt.Overwrite {
			if _, err := os.Stat(mp); err == nil {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrConflict,
					Message: "manifest file already exists",
					Op:      "olextract.write_artifacts",
					Path:    mp,
					Offset:  -1,
				}
			} else if !errors.Is(err, os.ErrNotExist) {
				return WriteResult{}, &olecfb.OLEError{
					Code:    olecfb.ErrUnsupported,
					Message: "check manifest file failed",
					Op:      "olextract.write_artifacts",
					Path:    mp,
					Offset:  -1,
					Cause:   err,
				}
			}
		}
		if err := writeManifestFile(mp, report, out); err != nil {
			return WriteResult{}, &olecfb.OLEError{
				Code:    olecfb.ErrUnsupported,
				Message: "write manifest failed",
				Op:      "olextract.write_artifacts",
				Path:    mp,
				Offset:  -1,
				Cause:   err,
			}
		}
		out.ManifestPath = mp
	}
	return out, nil
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

func buildArtifactFileName(index int, a olecfb.Artifact) string {
	base := sanitizeName(a.Path)
	if base == "" {
		base = "artifact"
	}
	if len(base) > 80 {
		base = base[:80]
	}
	ext := extensionByArtifact(a)
	return fmt.Sprintf("%06d_%s%s", index, base, ext)
}

func buildArtifactTreePath(index int, a olecfb.Artifact) string {
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
	fileName := segments[len(segments)-1]
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
	return strings.Trim(b.String(), "_")
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
	return os.WriteFile(path, buf, 0o644)
}

func extensionByArtifact(a olecfb.Artifact) string {
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
