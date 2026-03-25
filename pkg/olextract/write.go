package olextract

import (
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
)

type WriteOptions struct {
	Overwrite bool
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
	Files        []WrittenFile
}

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

	out := WriteResult{}
	for i, a := range report.Artifacts {
		if len(a.Raw) == 0 {
			out.Skipped++
			continue
		}
		name := buildArtifactFileName(i, a)
		p := filepath.Join(dstDir, name)
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
