package olecfb

import (
	"bytes"
	"testing"
)

func FuzzExtract(f *testing.F) {
	f.Add([]byte("not-a-cfb"))
	if seed, ok := buildFuzzSeedCFB(); ok {
		f.Add(seed)
	}

	f.Fuzz(func(t *testing.T, data []byte) {
		const maxInput = 4 << 20
		if len(data) > maxInput {
			return
		}

		openOpt := OpenOptions{
			Mode:           Lenient,
			MaxObjectCount: 4096,
			MaxTotalBytes:  16 << 20,
			MaxStreamBytes: 8 << 20,
			MaxChainLength: 1 << 20,
			MaxRecursion:   256,
		}
		file, err := OpenBytes(data, openOpt)
		if err != nil {
			return
		}
		defer file.Close()

		rep, err := file.Extract(ExtractOptions{
			Mode:              Lenient,
			IncludeRaw:        true,
			DetectImages:      true,
			DetectOLEDS:       true,
			UnwrapOle10Native: true,
			Deduplicate:       true,
			Limits: ExtractLimits{
				MaxDepth:        16,
				MaxArtifacts:    2048,
				MaxTotalBytes:   16 << 20,
				MaxArtifactSize: 8 << 20,
			},
		})
		if err != nil {
			return
		}
		if rep == nil {
			t.Fatal("nil report")
		}
		if rep.Stats.ArtifactsTotal != len(rep.Artifacts) {
			t.Fatalf("artifacts total mismatch: stats=%d len=%d", rep.Stats.ArtifactsTotal, len(rep.Artifacts))
		}
	})
}

func buildFuzzSeedCFB() ([]byte, bool) {
	f, err := CreateInMemory(CreateOptions{})
	if err != nil {
		return nil, false
	}
	defer f.Close()
	tx, err := f.Begin(TxOptions{})
	if err != nil {
		return nil, false
	}
	if err := tx.CreateStorage("/Docs"); err != nil {
		return nil, false
	}
	if err := tx.PutStream("/Docs/A", bytes.NewReader([]byte("abc")), 3); err != nil {
		return nil, false
	}
	if _, err := tx.Commit(nil, CommitOptions{}); err != nil {
		return nil, false
	}
	buf, err := f.SnapshotBytes()
	if err != nil {
		return nil, false
	}
	return buf, true
}
