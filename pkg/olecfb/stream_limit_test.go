package olecfb

import "testing"

func TestOpenStream_MaxStreamBytesLimit_NormalStream(t *testing.T) {
	fileBytes, _ := buildValidV4FileWithSingleNormalStream()
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict, MaxStreamBytes: 128})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if _, err := f.OpenStream("/Blob"); err == nil {
		t.Fatal("expected stream size limit error")
	} else if !IsCode(err, ErrLimitExceeded) {
		t.Fatalf("expected ErrLimitExceeded, got %v", err)
	}
}

func TestOpenStream_MaxStreamBytesLimit_MiniStream(t *testing.T) {
	fileBytes, _ := buildValidFileWithMiniStream()
	f, err := OpenBytes(fileBytes, OpenOptions{Mode: Strict, MaxStreamBytes: 8})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if _, err := f.OpenStream("/Small"); err == nil {
		t.Fatal("expected stream size limit error")
	} else if !IsCode(err, ErrLimitExceeded) {
		t.Fatalf("expected ErrLimitExceeded, got %v", err)
	}
}

