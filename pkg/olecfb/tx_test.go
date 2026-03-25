package olecfb

import (
	"bytes"
	"io"
	"path/filepath"
	"testing"
)

func TestTxPutStreamCommitInMemory(t *testing.T) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.CreateStorage("/Docs"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	payload := []byte("hello-tx")
	if err := tx.PutStream("/Docs/Note", bytes.NewReader(payload), int64(len(payload))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}

	n, err := f.GetNodeByPath("/Docs/Note")
	if err != nil {
		t.Fatalf("GetNodeByPath returned error: %v", err)
	}
	if n.Type != NodeStream {
		t.Fatalf("unexpected node type: %v", n.Type)
	}
	sr, err := f.OpenStream("/Docs/Note")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	got := make([]byte, len(payload))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if string(got) != string(payload) {
		t.Fatalf("unexpected payload: %q", string(got))
	}

	snap, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	f2, err := OpenBytes(snap, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes(snapshot) returned error: %v", err)
	}
	sr2, err := f2.OpenStream("/Docs/Note")
	if err != nil {
		t.Fatalf("OpenStream(snapshot) returned error: %v", err)
	}
	defer sr2.Close()
	got2 := make([]byte, len(payload))
	if _, err := io.ReadFull(sr2, got2); err != nil {
		t.Fatalf("ReadFull(snapshot) returned error: %v", err)
	}
	if string(got2) != string(payload) {
		t.Fatalf("unexpected payload from snapshot: %q", string(got2))
	}
}

func TestTxRenameDeleteAndRevert(t *testing.T) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, _ := f.Begin(TxOptions{})
	_ = tx.CreateStorage("/Docs")
	_ = tx.PutStream("/Docs/Note", bytes.NewReader([]byte("v1")), 2)
	_, _ = tx.Commit(nil, CommitOptions{})

	// Revert should keep old state.
	tx2, _ := f.Begin(TxOptions{})
	if err := tx2.Rename("/Docs/Note", "/Docs/NewNote"); err != nil {
		t.Fatalf("Rename returned error: %v", err)
	}
	if err := tx2.Revert(); err != nil {
		t.Fatalf("Revert returned error: %v", err)
	}
	if _, err := f.GetNodeByPath("/Docs/Note"); err != nil {
		t.Fatalf("expected original node after revert: %v", err)
	}
	if _, err := f.GetNodeByPath("/Docs/NewNote"); err == nil {
		t.Fatal("unexpected renamed node after revert")
	}

	// Commit rename.
	tx3, _ := f.Begin(TxOptions{})
	if err := tx3.Rename("/Docs/Note", "/Docs/NewNote"); err != nil {
		t.Fatalf("Rename returned error: %v", err)
	}
	if _, err := tx3.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if _, err := f.GetNodeByPath("/Docs/Note"); err == nil {
		t.Fatal("old path should not exist after rename commit")
	}
	if _, err := f.GetNodeByPath("/Docs/NewNote"); err != nil {
		t.Fatalf("new path should exist after rename commit: %v", err)
	}

	// Commit delete.
	tx4, _ := f.Begin(TxOptions{})
	if err := tx4.Delete("/Docs"); err != nil {
		t.Fatalf("Delete returned error: %v", err)
	}
	if _, err := tx4.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if _, err := f.GetNodeByPath("/Docs"); err == nil {
		t.Fatal("storage should be deleted")
	}
}

func TestTxCommitFileBackend(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "doc.cfb")
	f, err := CreateFile(p, CreateOptions{})
	if err != nil {
		t.Fatalf("CreateFile returned error: %v", err)
	}
	defer f.Close()
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.CreateStorage("/Docs"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	payload := []byte("tiny")
	if err := tx.PutStream("/Docs/Small", bytes.NewReader(payload), int64(len(payload))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	res, err := tx.Commit(nil, CommitOptions{Sync: true})
	if err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if res.BackendKind != "file" {
		t.Fatalf("unexpected backend kind: %s", res.BackendKind)
	}

	reopened, err := OpenFile(p, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenFile returned error: %v", err)
	}
	defer reopened.Close()
	sr, err := reopened.OpenStream("/Docs/Small")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	got := make([]byte, len(payload))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if string(got) != string(payload) {
		t.Fatalf("unexpected payload: %q", string(got))
	}
}

func TestTxMaxObjectLimit(t *testing.T) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	f.opt.MaxObjectCount = 2 // root + one node
	tx, _ := f.Begin(TxOptions{})
	if err := tx.CreateStorage("/One"); err != nil {
		t.Fatalf("CreateStorage returned error: %v", err)
	}
	if err := tx.CreateStorage("/Two"); err == nil {
		t.Fatal("expected object limit exceeded")
	}
}

func TestTxRenameKeepsExistingStreamData(t *testing.T) {
	base, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, _ := base.Begin(TxOptions{})
	_ = tx.CreateStorage("/Docs")
	_ = tx.PutStream("/Docs/Note", bytes.NewReader([]byte("keep")), 4)
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("seed commit returned error: %v", err)
	}
	snap, err := base.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}

	f, err := OpenBytesRW(snap, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	tx2, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx2.Rename("/Docs/Note", "/Docs/Renamed"); err != nil {
		t.Fatalf("Rename returned error: %v", err)
	}
	if _, err := tx2.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}

	sr, err := f.OpenStream("/Docs/Renamed")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	got := make([]byte, 4)
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if string(got) != "keep" {
		t.Fatalf("unexpected payload after rename: %q", string(got))
	}
}

func TestTxSingleActiveTransaction(t *testing.T) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx1, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if _, err := f.Begin(TxOptions{}); err == nil {
		t.Fatal("expected conflict while tx1 is active")
	} else if !IsCode(err, ErrConflict) {
		t.Fatalf("expected ErrConflict, got %v", err)
	}
	if err := tx1.Revert(); err != nil {
		t.Fatalf("Revert returned error: %v", err)
	}
	if _, err := f.Begin(TxOptions{}); err != nil {
		t.Fatalf("Begin should succeed after revert: %v", err)
	}
}

func TestTxCommitLargeStreamWithIncrementalFallback(t *testing.T) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}

	// 8 MiB stream ensures FAT sectors exceed DIFAT header slots (109).
	payload := bytes.Repeat([]byte{0xAB}, 8*1024*1024)
	if err := tx.PutStream("/Big", bytes.NewReader(payload), int64(len(payload))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	res, err := tx.Commit(nil, CommitOptions{Strategy: Incremental})
	if err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if res.StrategyUsed != FullRewrite {
		t.Fatalf("expected fallback strategy FullRewrite, got %v", res.StrategyUsed)
	}

	snap, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	reopened, err := OpenBytes(snap, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if reopened.hdr == nil {
		t.Fatal("missing header after reopen")
	}
	if reopened.hdr.NumDIFATSectors == 0 {
		t.Fatal("expected extended DIFAT sectors for large stream")
	}

	sr, err := reopened.OpenStream("/Big")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	head := make([]byte, 64)
	if _, err := io.ReadFull(sr, head); err != nil {
		t.Fatalf("ReadFull(head) returned error: %v", err)
	}
	for i, b := range head {
		if b != 0xAB {
			t.Fatalf("unexpected head byte at %d: 0x%02X", i, b)
		}
	}
	if _, err := sr.Seek(-1, io.SeekEnd); err != nil {
		t.Fatalf("Seek returned error: %v", err)
	}
	tail := make([]byte, 1)
	if _, err := io.ReadFull(sr, tail); err != nil {
		t.Fatalf("ReadFull(tail) returned error: %v", err)
	}
	if tail[0] != 0xAB {
		t.Fatalf("unexpected tail byte: 0x%02X", tail[0])
	}
}

func TestTxCommitPreservesV4Geometry(t *testing.T) {
	base, _ := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytesRW(base, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutStream("/Blob", bytes.NewReader([]byte("v4-keep")), int64(len("v4-keep"))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}

	snap, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	reopened, err := OpenBytes(snap, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if reopened.hdr == nil {
		t.Fatal("missing header")
	}
	if reopened.hdr.MajorVersion != cfbMajorVersion4 {
		t.Fatalf("unexpected major version: %d", reopened.hdr.MajorVersion)
	}
	if reopened.hdr.SectorShift != cfbSectorShiftV4 {
		t.Fatalf("unexpected sector shift: %d", reopened.hdr.SectorShift)
	}
	sr, err := reopened.OpenStream("/Blob")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	got := make([]byte, len("v4-keep"))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if string(got) != "v4-keep" {
		t.Fatalf("unexpected payload: %q", string(got))
	}
}

func TestTxCommitIncrementalInPlaceUpdate(t *testing.T) {
	base, _ := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytesRW(base, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	payload := bytes.Repeat([]byte{0x5A}, 4096)
	if err := tx.PutStream("/Blob", bytes.NewReader(payload), int64(len(payload))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	res, err := tx.Commit(nil, CommitOptions{Strategy: Incremental})
	if err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if res.StrategyUsed != Incremental {
		t.Fatalf("expected incremental strategy, got %v", res.StrategyUsed)
	}
	if res.BytesWritten != int64(len(payload)) {
		t.Fatalf("unexpected bytes written: got %d want %d", res.BytesWritten, len(payload))
	}
	if res.NewSize != int64(len(base)) {
		t.Fatalf("unexpected new size: got %d want %d", res.NewSize, len(base))
	}

	sr, err := f.OpenStream("/Blob")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	head := make([]byte, 8)
	if _, err := io.ReadFull(sr, head); err != nil {
		t.Fatalf("ReadFull(head) returned error: %v", err)
	}
	for i, b := range head {
		if b != 0x5A {
			t.Fatalf("unexpected head byte at %d: 0x%02X", i, b)
		}
	}
	if _, err := sr.Seek(-1, io.SeekEnd); err != nil {
		t.Fatalf("Seek returned error: %v", err)
	}
	tail := make([]byte, 1)
	if _, err := io.ReadFull(sr, tail); err != nil {
		t.Fatalf("ReadFull(tail) returned error: %v", err)
	}
	if tail[0] != 0x5A {
		t.Fatalf("unexpected tail byte: 0x%02X", tail[0])
	}
}

func TestTxCommitIncrementalMiniStreamUpdate(t *testing.T) {
	base, oldPayload := buildValidFileWithMiniStream()
	f, err := OpenBytesRW(base, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytesRW returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	payload := bytes.Repeat([]byte{0x4D}, len(oldPayload))
	if err := tx.PutStream("/Small", bytes.NewReader(payload), int64(len(payload))); err != nil {
		t.Fatalf("PutStream returned error: %v", err)
	}
	res, err := tx.Commit(nil, CommitOptions{Strategy: Incremental})
	if err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if res.StrategyUsed != Incremental {
		t.Fatalf("expected incremental strategy, got %v", res.StrategyUsed)
	}
	if res.BytesWritten != int64(len(payload)) {
		t.Fatalf("unexpected bytes written: got %d want %d", res.BytesWritten, len(payload))
	}

	sr, err := f.OpenStream("/Small")
	if err != nil {
		t.Fatalf("OpenStream returned error: %v", err)
	}
	defer sr.Close()
	got := make([]byte, len(payload))
	if _, err := io.ReadFull(sr, got); err != nil {
		t.Fatalf("ReadFull returned error: %v", err)
	}
	if !bytes.Equal(got, payload) {
		t.Fatal("unexpected mini stream payload after incremental commit")
	}

	snap, err := f.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	reopened, err := OpenBytes(snap, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	sr2, err := reopened.OpenStream("/Small")
	if err != nil {
		t.Fatalf("OpenStream(reopened) returned error: %v", err)
	}
	defer sr2.Close()
	got2 := make([]byte, len(payload))
	if _, err := io.ReadFull(sr2, got2); err != nil {
		t.Fatalf("ReadFull(reopened) returned error: %v", err)
	}
	if !bytes.Equal(got2, payload) {
		t.Fatal("unexpected mini stream payload after reopen")
	}
}

func TestTxCommitIncrementalMultipleStreamsFallback(t *testing.T) {
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
		t.Fatalf("seed commit returned error: %v", err)
	}

	tx2, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx2.PutStream("/A", bytes.NewReader([]byte("aaaa")), 4); err != nil {
		t.Fatalf("PutStream /A returned error: %v", err)
	}
	if err := tx2.PutStream("/B", bytes.NewReader([]byte("bbbb")), 4); err != nil {
		t.Fatalf("PutStream /B returned error: %v", err)
	}
	res, err := tx2.Commit(nil, CommitOptions{Strategy: Incremental})
	if err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	if res.StrategyUsed != FullRewrite {
		t.Fatalf("expected fallback FullRewrite, got %v", res.StrategyUsed)
	}

	assertStreamEquals(t, f, "/A", "aaaa")
	assertStreamEquals(t, f, "/B", "bbbb")
}

func assertStreamEquals(t *testing.T, f *File, path, want string) {
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
