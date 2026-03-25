package olecfb

import (
	"bytes"
	"testing"
)

func TestExtract_DeduplicateAcrossStreams(t *testing.T) {
	f := buildExtractTestFile(t)
	rep, err := f.Extract(ExtractOptions{Deduplicate: true})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	// /A and /B share payload; dedup should keep one and count one dedup hit.
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
	if rep.Stats.DedupHits != 1 {
		t.Fatalf("unexpected dedup hits: %d", rep.Stats.DedupHits)
	}
}

func TestExtract_MaxTotalBytesLimit(t *testing.T) {
	f := buildExtractTestFile(t)
	rep, err := f.Extract(ExtractOptions{
		Deduplicate: false,
		Limits: ExtractLimits{
			MaxTotalBytes: 7, // one stream fits (4), second would overflow.
		},
	})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if !rep.Partial {
		t.Fatal("expected partial report due to total bytes limit")
	}
	if rep.Stats.BytesExported > 4 {
		t.Fatalf("unexpected exported bytes: %d", rep.Stats.BytesExported)
	}
	foundQuota := false
	for _, w := range rep.Warnings {
		if w.Code == ErrQuotaExceeded {
			foundQuota = true
			break
		}
	}
	if !foundQuota {
		t.Fatal("expected quota exceeded warning")
	}
}

func TestExtract_MaxArtifactsLimit(t *testing.T) {
	f := buildExtractTestFile(t)
	rep, err := f.Extract(ExtractOptions{
		Deduplicate: false,
		Limits: ExtractLimits{
			MaxArtifacts: 1,
		},
	})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if !rep.Partial {
		t.Fatal("expected partial report due to artifact limit")
	}
	if rep.Stats.ArtifactsTotal != 1 {
		t.Fatalf("unexpected artifacts total: %d", rep.Stats.ArtifactsTotal)
	}
	foundLimit := false
	for _, w := range rep.Warnings {
		if w.Code == ErrLimitExceeded {
			foundLimit = true
			break
		}
	}
	if !foundLimit {
		t.Fatal("expected limit exceeded warning")
	}
}

func TestExtract_DeterministicOrder(t *testing.T) {
	f := buildExtractTestFile(t)
	r1, err := f.Extract(ExtractOptions{Deduplicate: false})
	if err != nil {
		t.Fatalf("Extract #1 returned error: %v", err)
	}
	r2, err := f.Extract(ExtractOptions{Deduplicate: false})
	if err != nil {
		t.Fatalf("Extract #2 returned error: %v", err)
	}
	if len(r1.Artifacts) != len(r2.Artifacts) {
		t.Fatalf("artifact count mismatch: %d vs %d", len(r1.Artifacts), len(r2.Artifacts))
	}
	for i := range r1.Artifacts {
		a, b := r1.Artifacts[i], r2.Artifacts[i]
		if a.Path != b.Path || a.SHA256 != b.SHA256 || a.Kind != b.Kind || a.ParentID != b.ParentID || a.Depth != b.Depth {
			t.Fatalf("artifact mismatch at index %d", i)
		}
	}
}

func TestExtract_UsesOpenOptionDefaultStreamLimit(t *testing.T) {
	base := buildExtractTestFile(t)
	snap, err := base.SnapshotBytes()
	if err != nil {
		t.Fatalf("SnapshotBytes returned error: %v", err)
	}
	f, err := OpenBytes(snap, OpenOptions{
		Mode:           Strict,
		MaxStreamBytes: 3, // each stream is 4 bytes; extract should skip due inherited artifact limit.
	})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	rep, err := f.Extract(ExtractOptions{Deduplicate: false})
	if err != nil {
		t.Fatalf("Extract returned error: %v", err)
	}
	if !rep.Partial {
		t.Fatal("expected partial report due to inherited max stream bytes")
	}
	if rep.Stats.ArtifactsTotal != 0 {
		t.Fatalf("expected zero artifacts due to size skip, got %d", rep.Stats.ArtifactsTotal)
	}
	foundLimit := false
	for _, w := range rep.Warnings {
		if w.Code == ErrLimitExceeded {
			foundLimit = true
			break
		}
	}
	if !foundLimit {
		t.Fatal("expected limit exceeded warning")
	}
}

func buildExtractTestFile(t *testing.T) *File {
	t.Helper()
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		t.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		t.Fatalf("Begin returned error: %v", err)
	}
	if err := tx.PutStream("/A", bytes.NewReader([]byte("same")), 4); err != nil {
		t.Fatalf("PutStream /A returned error: %v", err)
	}
	if err := tx.PutStream("/B", bytes.NewReader([]byte("same")), 4); err != nil {
		t.Fatalf("PutStream /B returned error: %v", err)
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		t.Fatalf("Commit returned error: %v", err)
	}
	return f
}
