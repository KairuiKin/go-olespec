package olecfb

import "testing"

func TestWalkFATChain_OK(t *testing.T) {
	fat := []uint32{
		1,
		2,
		cfbEndOfChain,
	}
	chain, err := walkFATChain(fat, 0, 0)
	if err != nil {
		t.Fatalf("walkFATChain returned error: %v", err)
	}
	if len(chain) != 3 {
		t.Fatalf("unexpected chain length: %d", len(chain))
	}
	if chain[0] != 0 || chain[1] != 1 || chain[2] != 2 {
		t.Fatalf("unexpected chain values: %#v", chain)
	}
}

func TestWalkFATChain_Cycle(t *testing.T) {
	fat := []uint32{
		1,
		0,
	}
	_, err := walkFATChain(fat, 0, 0)
	if err == nil {
		t.Fatal("expected cycle error")
	}
	if !IsCode(err, ErrCycleDetected) {
		t.Fatalf("expected ErrCycleDetected, got %v", err)
	}
}

func TestWalkFATChain_OutOfBounds(t *testing.T) {
	fat := []uint32{
		10,
	}
	_, err := walkFATChain(fat, 0, 0)
	if err == nil {
		t.Fatal("expected out-of-bounds error")
	}
	if !IsCode(err, ErrOutOfBounds) {
		t.Fatalf("expected ErrOutOfBounds, got %v", err)
	}
}

func TestWalkFATChain_LimitExceeded(t *testing.T) {
	fat := []uint32{
		1,
		2,
		cfbEndOfChain,
	}
	_, err := walkFATChain(fat, 0, 2)
	if err == nil {
		t.Fatal("expected limit error")
	}
	if !IsCode(err, ErrLimitExceeded) {
		t.Fatalf("expected ErrLimitExceeded, got %v", err)
	}
}
