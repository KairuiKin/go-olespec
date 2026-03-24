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
	if _, err := tx.Commit(nil, CommitOptions{Strategy: Incremental}); err != nil {
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
