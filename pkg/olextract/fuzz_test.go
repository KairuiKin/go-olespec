package olextract

import (
	"errors"
	"strings"
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
)

func FuzzWriteArtifacts(f *testing.F) {
	f.Add([]byte("abc"), "/Docs/A", "hello.txt", "manifest.json", uint8(0))
	f.Add([]byte{}, "/Docs/B", "", "", uint8(1))
	f.Add([]byte("payload"), "/Embedded!/Blob", "C:/tmp/final.bin", "../bad.json", uint8(2))

	f.Fuzz(func(t *testing.T, raw []byte, artPath, oleFileName, manifestName string, mode uint8) {
		const (
			maxRaw   = 2 << 20
			maxField = 256
		)
		if len(raw) > maxRaw {
			return
		}
		artPath = truncateUTF8(artPath, maxField)
		oleFileName = truncateUTF8(oleFileName, maxField)
		manifestName = truncateUTF8(manifestName, maxField)

		kind := olecfb.ArtifactStream
		if mode&0x01 != 0 {
			kind = olecfb.ArtifactOleObj
		}
		layout := WriteLayoutFlat
		if mode&0x02 != 0 {
			layout = WriteLayoutTree
		}
		writeManifest := mode&0x04 != 0
		atomic := mode&0x08 != 0
		overwrite := mode&0x10 != 0
		preferOLEFileName := mode&0x20 != 0

		rep := &olecfb.ExtractReport{
			Artifacts: []olecfb.Artifact{
				{
					Path:        artPath,
					Kind:        kind,
					OLEFileName: oleFileName,
					Raw:         append([]byte(nil), raw...),
				},
				{
					Path: "/skip/no-raw",
					Kind: olecfb.ArtifactStream,
				},
			},
		}

		opt := WriteOptions{
			Layout:            layout,
			WriteManifest:     writeManifest,
			ManifestName:      manifestName,
			AtomicPublish:     atomic,
			Overwrite:         overwrite,
			PreferOLEFileName: preferOLEFileName,
		}
		res, err := WriteArtifacts(rep, t.TempDir(), opt)
		if err != nil {
			var oe *olecfb.OLEError
			if !errors.As(err, &oe) {
				t.Fatalf("expected *OLEError, got %T", err)
			}
			return
		}

		if len(raw) == 0 {
			if res.FilesWritten != 0 {
				t.Fatalf("expected no written files for empty raw, got %d", res.FilesWritten)
			}
		} else if res.FilesWritten != 1 {
			t.Fatalf("expected one written file, got %d", res.FilesWritten)
		}
		if writeManifest && res.ManifestPath == "" {
			t.Fatal("manifest path should not be empty when write_manifest succeeds")
		}
	})
}

func truncateUTF8(v string, n int) string {
	if n <= 0 || len(v) <= n {
		return v
	}
	var b strings.Builder
	b.Grow(n)
	for _, r := range v {
		next := b.Len() + len(string(r))
		if next > n {
			break
		}
		b.WriteRune(r)
	}
	return b.String()
}
