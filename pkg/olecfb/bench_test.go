package olecfb

import (
	"bytes"
	"fmt"
	"testing"
)

func BenchmarkExtractFlat(b *testing.B) {
	f := buildBenchExtractFile(b, 64, 1024)
	opt := ExtractOptions{
		Deduplicate:  false,
		DetectImages: true,
		DetectOLEDS:  true,
		Limits: ExtractLimits{
			MaxDepth:      2,
			MaxArtifacts:  4096,
			MaxTotalBytes: 1 << 30,
		},
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		rep, err := f.Extract(opt)
		if err != nil {
			b.Fatalf("Extract returned error: %v", err)
		}
		if rep.Stats.ArtifactsTotal != 64 {
			b.Fatalf("unexpected artifact total: %d", rep.Stats.ArtifactsTotal)
		}
	}
}

func BenchmarkExtractRecursive(b *testing.B) {
	innerBytes, _ := buildValidV4FileWithSingleNormalStream()
	midBytes := buildValidV4FileWithBigNamedStream("InnerOLE", innerBytes)
	outerBytes := buildValidV4FileWithBigNamedStream("Embedded", midBytes)
	f, err := OpenBytes(outerBytes, OpenOptions{Mode: Strict})
	if err != nil {
		b.Fatalf("OpenBytes returned error: %v", err)
	}

	opt := ExtractOptions{
		Deduplicate: false,
		Limits: ExtractLimits{
			MaxDepth:      4,
			MaxArtifacts:  128,
			MaxTotalBytes: 1 << 30,
		},
	}

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		rep, err := f.Extract(opt)
		if err != nil {
			b.Fatalf("Extract returned error: %v", err)
		}
		if rep.Stats.ArtifactsTotal != 3 {
			b.Fatalf("unexpected artifact total: %d", rep.Stats.ArtifactsTotal)
		}
	}
}

func BenchmarkCommitFullRewrite(b *testing.B) {
	base, _ := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytesRW(base, OpenOptions{Mode: Strict})
	if err != nil {
		b.Fatalf("OpenBytesRW returned error: %v", err)
	}
	payload := bytes.Repeat([]byte{0x61}, 4096)

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		payload[0] = byte(i)
		tx, err := f.Begin(TxOptions{})
		if err != nil {
			b.Fatalf("Begin returned error: %v", err)
		}
		if err := tx.PutStream("/Blob", bytes.NewReader(payload), int64(len(payload))); err != nil {
			b.Fatalf("PutStream returned error: %v", err)
		}
		res, err := tx.Commit(nil, CommitOptions{Strategy: FullRewrite})
		if err != nil {
			b.Fatalf("Commit returned error: %v", err)
		}
		if res.StrategyUsed != FullRewrite {
			b.Fatalf("unexpected strategy used: %v", res.StrategyUsed)
		}
	}
}

func BenchmarkCommitIncrementalInPlace(b *testing.B) {
	base, _ := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytesRW(base, OpenOptions{Mode: Strict})
	if err != nil {
		b.Fatalf("OpenBytesRW returned error: %v", err)
	}
	payload := bytes.Repeat([]byte{0x5A}, 4096)

	b.ReportAllocs()
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		payload[0] = byte(i)
		tx, err := f.Begin(TxOptions{})
		if err != nil {
			b.Fatalf("Begin returned error: %v", err)
		}
		if err := tx.PutStream("/Blob", bytes.NewReader(payload), int64(len(payload))); err != nil {
			b.Fatalf("PutStream returned error: %v", err)
		}
		res, err := tx.Commit(nil, CommitOptions{Strategy: Incremental})
		if err != nil {
			b.Fatalf("Commit returned error: %v", err)
		}
		if res.StrategyUsed != Incremental {
			b.Fatalf("unexpected strategy used: %v", res.StrategyUsed)
		}
	}
}

func buildBenchExtractFile(b *testing.B, streamCount, streamSize int) *File {
	b.Helper()
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		b.Fatalf("CreateInMemory returned error: %v", err)
	}
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		b.Fatalf("Begin returned error: %v", err)
	}
	for i := 0; i < streamCount; i++ {
		path := fmt.Sprintf("/S%03d", i)
		fill := byte(0x41 + (i % 26))
		payload := bytes.Repeat([]byte{fill}, streamSize)
		if err := tx.PutStream(path, bytes.NewReader(payload), int64(streamSize)); err != nil {
			b.Fatalf("PutStream(%s) returned error: %v", path, err)
		}
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		b.Fatalf("Commit returned error: %v", err)
	}
	return f
}
