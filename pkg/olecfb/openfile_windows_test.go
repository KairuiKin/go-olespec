package olecfb

import (
	"os"
	"path/filepath"
	"runtime"
	"testing"
)

func TestOpenFileFailureClosesHandle(t *testing.T) {
	if runtime.GOOS != "windows" {
		t.Skip("windows-specific handle semantics")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "bad.cfb")
	if err := os.WriteFile(p, []byte{0x01, 0x02, 0x03}, 0o644); err != nil {
		t.Fatalf("WriteFile returned error: %v", err)
	}
	if _, err := OpenFile(p, OpenOptions{Mode: Strict}); err == nil {
		t.Fatal("expected parse error for invalid cfb")
	}
	// If OpenFile leaked the handle, Remove may fail on Windows.
	if err := os.Remove(p); err != nil {
		t.Fatalf("Remove returned error: %v", err)
	}
}

