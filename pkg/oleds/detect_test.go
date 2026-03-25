package oleds

import (
	"bytes"
	"encoding/binary"
	"testing"
)

func TestDetectOle10NativeByPath(t *testing.T) {
	data := buildOle10Native("hello.txt", "C:\\tmp\\hello.txt", []byte("hello"))
	d := Detect("/Object/\x01Ole10Native", data)
	if d.Kind != KindOle10Native {
		t.Fatalf("unexpected kind: %s", d.Kind)
	}
	if d.FileName != "hello.txt" {
		t.Fatalf("unexpected file name: %s", d.FileName)
	}
	if d.PayloadSize != 5 {
		t.Fatalf("unexpected payload size: %d", d.PayloadSize)
	}
}

func TestDetectCompObjByPath(t *testing.T) {
	data := buildCompObj("Paintbrush Picture", "Paint.Picture")
	d := Detect("/Object/CompObj", data)
	if d.Kind != KindCompObj {
		t.Fatalf("unexpected kind: %s", d.Kind)
	}
	if d.ProgID != "Paint.Picture" {
		t.Fatalf("unexpected progid: %s", d.ProgID)
	}
}

func TestDetectPackageByPath(t *testing.T) {
	d := Detect("/Object/Package", []byte{1, 2, 3})
	if d.Kind != KindPackage {
		t.Fatalf("unexpected kind: %s", d.Kind)
	}
}

func TestParseOle10Native(t *testing.T) {
	data := buildOle10Native("hello.txt", "C:\\tmp\\hello.txt", []byte("hello"))
	n, ok := ParseOle10Native(data)
	if !ok {
		t.Fatal("expected ParseOle10Native success")
	}
	if n.FileName != "hello.txt" {
		t.Fatalf("unexpected file name: %q", n.FileName)
	}
	if n.SourcePath != "C:\\tmp\\hello.txt" {
		t.Fatalf("unexpected source path: %q", n.SourcePath)
	}
	if string(n.Payload) != "hello" {
		t.Fatalf("unexpected payload: %q", string(n.Payload))
	}
}

func TestParseOle10NativeInvalid(t *testing.T) {
	if _, ok := ParseOle10Native([]byte("broken")); ok {
		t.Fatal("expected ParseOle10Native failure")
	}
}

func buildOle10Native(fileName, sourcePath string, payload []byte) []byte {
	var body bytes.Buffer
	body.Write([]byte{0x02, 0x00}) // unknown short
	body.WriteString(fileName)
	body.WriteByte(0)
	body.WriteString(sourcePath)
	body.WriteByte(0)
	body.WriteString(sourcePath)
	body.WriteByte(0)
	_ = binary.Write(&body, binary.LittleEndian, uint32(len(payload)))
	body.Write(payload)

	out := make([]byte, 4)
	binary.LittleEndian.PutUint32(out, uint32(body.Len()))
	out = append(out, body.Bytes()...)
	return out
}

func buildCompObj(userType, progID string) []byte {
	data := make([]byte, 28)
	data = append(data, []byte(userType)...)
	data = append(data, 0)
	data = append(data, []byte("CLIPFORMAT")...)
	data = append(data, 0)
	data = append(data, []byte(progID)...)
	data = append(data, 0)
	return data
}
