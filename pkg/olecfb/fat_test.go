package olecfb

import (
	"encoding/binary"
	"testing"
)

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

func TestOpenBytes_LoadFATExtendedDIFAT(t *testing.T) {
	const fatCount = 110
	header := buildValidHeader(cfbMajorVersion3)
	binary.LittleEndian.PutUint32(header[44:48], fatCount)
	binary.LittleEndian.PutUint32(header[68:72], 0) // first DIFAT sector
	binary.LittleEndian.PutUint32(header[72:76], 1) // one DIFAT sector
	for i := 0; i < cfbNumDifatEntries; i++ {
		v := uint32(cfbFreeSector)
		if i < cfbNumDifatEntries {
			v = uint32(i + 1) // FAT sectors 1..109
		}
		binary.LittleEndian.PutUint32(header[76+i*4:80+i*4], v)
	}

	totalSectors := 111 // 1 DIFAT + 110 FAT
	buf := make([]byte, cfbHeaderSize+totalSectors*cfbHeaderSize)
	copy(buf, header)

	// DIFAT sector(0): one extra FAT sector id (110), then free entries, then EOC.
	difatOff := cfbHeaderSize
	binary.LittleEndian.PutUint32(buf[difatOff:difatOff+4], 110)
	for i := 1; i < cfbHeaderSize/4-1; i++ {
		binary.LittleEndian.PutUint32(buf[difatOff+i*4:difatOff+i*4+4], cfbFreeSector)
	}
	binary.LittleEndian.PutUint32(buf[difatOff+(cfbHeaderSize/4-1)*4:difatOff+(cfbHeaderSize/4-1)*4+4], cfbEndOfChain)

	f, err := OpenBytes(buf, OpenOptions{Mode: Strict})
	if err != nil {
		t.Fatalf("OpenBytes returned error: %v", err)
	}
	if got := len(f.fat); got != fatCount*(cfbHeaderSize/4) {
		t.Fatalf("unexpected fat length: got %d want %d", got, fatCount*(cfbHeaderSize/4))
	}
}
