package olextract

import (
	"bytes"
	"errors"
	"io"
	"os"
	"path/filepath"
	"testing"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
	"github.com/KairuiKin/go-olespec/pkg/olecfb/storage"
)

func TestExtractBytes(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractBytes(
		buf,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractBytes returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
	if rep.Artifacts[0].Path != "/Docs/A" {
		t.Fatalf("unexpected artifact path: %s", rep.Artifacts[0].Path)
	}
}

func TestExtractBackend(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rb := &testReadBackend{buf: buf}
	rep, err := ExtractBackend(
		rb,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractBackend returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
	if !rb.closed {
		t.Fatal("backend should be closed")
	}
}

func TestExtractBackendNil(t *testing.T) {
	if _, err := ExtractBackend(nil, olecfb.OpenOptions{}, olecfb.ExtractOptions{}); err == nil {
		t.Fatal("expected error for nil backend")
	} else if !olecfb.IsCode(err, olecfb.ErrInvalidArgument) {
		t.Fatalf("expected ErrInvalidArgument, got %v", err)
	}
}

func TestExtractBackendOpenFailureClosesBackend(t *testing.T) {
	rb := &testReadBackend{buf: []byte("not-cfb")}
	if _, err := ExtractBackend(rb, olecfb.OpenOptions{Mode: olecfb.Strict}, olecfb.ExtractOptions{}); err == nil {
		t.Fatal("expected open failure")
	}
	if !rb.closed {
		t.Fatal("backend should be closed on open failure")
	}
}

func TestExtractFile(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	dir := t.TempDir()
	p := filepath.Join(dir, "sample.cfb")
	if err := os.WriteFile(p, buf, 0o644); err != nil {
		t.Fatalf("WriteFile returned error: %v", err)
	}
	rep, err := ExtractFile(
		p,
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractFile returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
}

func TestExtractReader(t *testing.T) {
	buf := buildSampleCFBBytes(t)
	rep, err := ExtractReader(
		bytes.NewReader(buf),
		olecfb.OpenOptions{Mode: olecfb.Strict},
		olecfb.ExtractOptions{Deduplicate: true},
	)
	if err != nil {
		t.Fatalf("ExtractReader returned error: %v", err)
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
}

func TestExtractReaderNil(t *testing.T) {
	if _, err := ExtractReader(nil, olecfb.OpenOptions{}, olecfb.ExtractOptions{}); err == nil {
		t.Fatal("expected error for nil reader")
	} else if !olecfb.IsCode(err, olecfb.ErrInvalidArgument) {
		t.Fatalf("expected ErrInvalidArgument, got %v", err)
	}
}

func TestExtractReaderReadFailure(t *testing.T) {
	if _, err := ExtractReader(failReader{}, olecfb.OpenOptions{}, olecfb.ExtractOptions{}); err == nil {
		t.Fatal("expected read failure")
	} else if !olecfb.IsCode(err, olecfb.ErrUnsupported) {
		t.Fatalf("expected ErrUnsupported, got %v", err)
	}
}

func buildSampleCFBBytes(t *testing.T) []byte {
	t.Helper()
	f, err := olecfb.CreateInMemory(olecfb.CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
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

type failReader struct{}

func (failReader) Read(_ []byte) (int, error) {
	return 0, errors.New("boom")
}

type testReadBackend struct {
	buf    []byte
	closed bool
}

var _ storage.ReadBackend = (*testReadBackend)(nil)

func (b *testReadBackend) ReadAt(p []byte, off int64) (int, error) {
	if off < 0 || off >= int64(len(b.buf)) {
		return 0, io.EOF
	}
	n := copy(p, b.buf[off:])
	if n < len(p) {
		return n, io.EOF
	}
	return n, nil
}

func (b *testReadBackend) Size() int64 { return int64(len(b.buf)) }

func (b *testReadBackend) Close() error {
	b.closed = true
	return nil
}
