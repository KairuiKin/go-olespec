package olecfb

import (
	"bytes"
	"errors"
	"io"
	"testing"
)

func TestTxCommitFailureDoesNotPolluteState(t *testing.T) {
	base := buildAtomicSeedBytes(t)
	backend := &failingWriteBackend{
		buf: append([]byte(nil), base...),
		err: errors.New("forced write failure"),
	}
	f, err := OpenReadWrite(backend, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenReadWrite returned error: %v", err)
	}

	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.CreateStorage("/Docs"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	if err := tx.PutStream("/Docs/New", bytes.NewReader([]byte("new-data")), int64(len("new-data"))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err == nil {
		t.Fatal("expected commit error")
	} else if !IsCode(err, ErrCommitFailed) {
		t.Fatalf("expected ErrCommitFailed, got %v", err)
	}

	if _, err := f.GetNodeByPath("/Docs/New"); err == nil {
		t.Fatal("new node should not exist after failed commit")
	}
	seedNode, err := f.GetNodeByPath("/Seed")
	if err != nil {
		t.Fatalf("seed node should remain: %v", err)
	}
	if seedNode.Type != NodeStream {
		t.Fatalf("unexpected seed node type: %v", seedNode.Type)
	}
	sr, err := f.OpenStream("/Seed")
	if err != nil {
		t.Fatalf("OpenStream(/Seed) returned error: %v", err)
	}
	defer sr.Close()
	got := make([]byte, len("seed"))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if string(got) != "seed" {
		t.Fatalf("unexpected seed payload: %q", string(got))
	}

	// Failed commit should release active tx lock.
	if _, err := f.Begin(TxOptions{}); err != nil {
		t.Fatalf("Begin after failed commit should succeed: %v", err)
	}
}

func TestTxCommitIncrementalFallbackFailureDoesNotPolluteState(t *testing.T) {
	base := buildAtomicTwoStreamSeedBytes(t)
	backend := &failingWriteBackend{
		buf: append([]byte(nil), base...),
		err: errors.New("forced write failure"),
	}
	f, err := OpenReadWrite(backend, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenReadWrite returned error: %v", err)
	}

	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutStream("/A", bytes.NewReader([]byte("aaaa")), 4); err != nil {
		t.Fatalf("PutStream /A returned error: %v", err)
	}
	if err := tx.PutStream("/B", bytes.NewReader([]byte("bbbb")), 4); err != nil {
		t.Fatalf("PutStream /B returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{Strategy: Incremental}); err == nil {
		t.Fatal("expected commit error")
	} else if !IsCode(err, ErrCommitFailed) {
		t.Fatalf("expected ErrCommitFailed, got %v", err)
	}

	assertAtomicStreamEquals(t, f, "/A", "1111")
	assertAtomicStreamEquals(t, f, "/B", "2222")

	if _, err := f.Begin(TxOptions{}); err != nil {
		t.Fatalf("Begin after failed fallback commit should succeed: %v", err)
	}
}

func buildAtomicSeedBytes(t *testing.T) []byte {
	t.Helper()
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutStream("/Seed", bytes.NewReader([]byte("seed")), 4); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	buf, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	return buf
}

func buildAtomicTwoStreamSeedBytes(t *testing.T) []byte {
	t.Helper()
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutStream("/A", bytes.NewReader([]byte("1111")), 4); err != nil {
		t.Fatalf("PutStream /A returned error: %v", err)
	}
	if err := tx.PutStream("/B", bytes.NewReader([]byte("2222")), 4); err != nil {
		t.Fatalf("PutStream /B returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	buf, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	return buf
}

func assertAtomicStreamEquals(t *testing.T, f *File, path, want string) {
	t.Helper()
	sr, err := f.OpenStream(path)
	if err != nil {
		t.Fatalf("OpenStream(%s) returned error: %v", path, err)
	}
	defer sr.Close()
	got := make([]byte, len(want))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull(%s) returned error: %v", path, err)
	}
	if string(got) != want {
		t.Fatalf("unexpected payload for %s: got %q want %q", path, string(got), want)
	}
}

type failingWriteBackend struct {
	buf []byte
	err error
}

func (b *failingWriteBackend) ReadAt(p []byte, off int64) (int, error) {
	if off < 0 {
		return 0, io.EOF
	}
	if off >= int64(len(b.buf)) {
		return 0, io.EOF
	}
	n := copy(p, b.buf[off:])
	if n < len(p) {
		return n, io.EOF
	}
	return n, nil
}

func (b *failingWriteBackend) Size() int64 { return int64(len(b.buf)) }
func (b *failingWriteBackend) Close() error { return nil }
func (b *failingWriteBackend) Sync() error  { return nil }

func (b *failingWriteBackend) WriteAt(p []byte, off int64) (int, error) {
	return 0, b.err
}

func (b *failingWriteBackend) Truncate(size int64) error { return nil }
