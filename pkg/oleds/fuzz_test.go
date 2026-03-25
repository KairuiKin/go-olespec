package oleds

import (
	"bytes"
	"encoding/binary"
	"testing"
)

func FuzzParseOle10Native(f *testing.F) {
	f.Add([]byte("x"))
	f.Add(buildValidOle10NativeSeed("sample.txt", "C:\\src\\sample.txt", "C:\\tmp\\sample.txt", []byte("hello")))

	f.Fuzz(func(t *testing.T, data []byte) {
		const maxInput = 2 << 20
		if len(data) > maxInput {
			return
		}

		_ = Detect("\x01Ole10Native", data)
		_ = Detect("\x01CompObj", data)
		_ = Detect("Package", data)

		native, ok := ParseOle10Native(data)
		if !ok {
			return
		}
		if len(native.Payload) > len(data) {
			t.Fatalf("payload too large: payload=%d data=%d", len(native.Payload), len(data))
		}

		// ParseOle10Native is expected to copy payload bytes.
		mut := append([]byte(nil), data...)
		native2, ok2 := ParseOle10Native(mut)
		if !ok2 || len(native2.Payload) == 0 {
			return
		}
		before := native2.Payload[0]
		for i := range mut {
			mut[i] ^= 0xFF
		}
		if native2.Payload[0] != before {
			t.Fatal("payload aliases input buffer")
		}
	})
}

func buildValidOle10NativeSeed(fileName, sourcePath, tempPath string, payload []byte) []byte {
	body := make([]byte, 0, 64+len(payload))
	body = append(body, 0x00, 0x00)
	body = append(body, []byte(fileName)...)
	body = append(body, 0)
	body = append(body, []byte(sourcePath)...)
	body = append(body, 0)
	body = append(body, []byte(tempPath)...)
	body = append(body, 0)
	var size [4]byte
	binary.LittleEndian.PutUint32(size[:], uint32(len(payload)))
	body = append(body, size[:]...)
	body = append(body, payload...)

	total := uint32(len(body))
	out := make([]byte, 4)
	binary.LittleEndian.PutUint32(out, total)
	out = append(out, body...)
	return bytes.Clone(out)
}
