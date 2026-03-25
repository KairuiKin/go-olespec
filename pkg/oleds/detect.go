package oleds

import (
	"bytes"
	"encoding/binary"
	"path"
	"strings"
)

type Kind string

const (
	KindUnknown     Kind = "unknown"
	KindOle10Native Kind = "ole10native"
	KindCompObj     Kind = "compobj"
	KindPackage     Kind = "package"
)

type Detection struct {
	Kind        Kind
	FileName    string
	SourcePath  string
	ProgID      string
	UserType    string
	PayloadSize uint32
	Confidence  float64
}

type Ole10Native struct {
	FileName   string
	SourcePath string
	TempPath   string
	Payload    []byte
}

func Detect(streamPath string, data []byte) Detection {
	base := strings.ToLower(path.Base(streamPath))
	switch base {
	case "\x01ole10native":
		if d, ok := parseOle10Native(data); ok {
			d.Confidence = 0.95
			return d
		}
		return Detection{Kind: KindOle10Native, Confidence: 0.5}
	case "\x01compobj", "compobj":
		if d, ok := parseCompObj(data); ok {
			d.Confidence = 0.85
			return d
		}
		return Detection{Kind: KindCompObj, Confidence: 0.4}
	case "package":
		return Detection{Kind: KindPackage, Confidence: 0.9}
	}

	if d, ok := parseOle10Native(data); ok {
		d.Confidence = 0.8
		return d
	}
	if d, ok := parseCompObj(data); ok {
		d.Confidence = 0.65
		return d
	}
	if bytes.HasPrefix(data, []byte("PK\x03\x04")) {
		return Detection{Kind: KindPackage, Confidence: 0.4}
	}
	return Detection{Kind: KindUnknown}
}

func parseOle10Native(data []byte) (Detection, bool) {
	native, ok := parseOle10NativePayload(data, false)
	if !ok {
		return Detection{}, false
	}
	return Detection{
		Kind:        KindOle10Native,
		FileName:    native.FileName,
		SourcePath:  native.SourcePath,
		PayloadSize: uint32(len(native.Payload)),
	}, true
}

// ParseOle10Native parses an Ole10Native stream and returns metadata with payload.
func ParseOle10Native(data []byte) (Ole10Native, bool) {
	return parseOle10NativePayload(data, true)
}

func parseOle10NativePayload(data []byte, copyPayload bool) (Ole10Native, bool) {
	// Heuristic parser for Ole10Native stream:
	// [0:4] total size, [4:6] unknown, then 3 ANSI C strings, then [4] payload size, payload bytes.
	if len(data) < 16 {
		return Ole10Native{}, false
	}
	total := binary.LittleEndian.Uint32(data[0:4])
	if total == 0 {
		return Ole10Native{}, false
	}
	off := 6
	fileName, n, ok := readCString(data, off)
	if !ok {
		return Ole10Native{}, false
	}
	off += n
	sourcePath, n, ok := readCString(data, off)
	if !ok {
		return Ole10Native{}, false
	}
	off += n
	tempPath, n, ok := readCString(data, off)
	if !ok {
		return Ole10Native{}, false
	}
	off += n
	if off+4 > len(data) {
		return Ole10Native{}, false
	}
	payloadSize := binary.LittleEndian.Uint32(data[off : off+4])
	off += 4
	if payloadSize > uint32(len(data)-off) {
		return Ole10Native{}, false
	}
	payload := data[off : off+int(payloadSize)]
	if copyPayload {
		payload = append([]byte(nil), payload...)
	}
	return Ole10Native{
		FileName:    fileName,
		SourcePath:  sourcePath,
		TempPath:    tempPath,
		Payload:     payload,
	}, true
}

func parseCompObj(data []byte) (Detection, bool) {
	// Best-effort parser: CompObj streams usually include a few ANSI zero-terminated strings
	// after the fixed-size header area.
	if len(data) < 28 {
		return Detection{}, false
	}
	off := 28
	userType, n, ok := readCString(data, off)
	if !ok {
		return Detection{}, false
	}
	off += n
	clipOrProg, n, ok := readCString(data, off)
	if !ok {
		return Detection{}, false
	}
	off += n
	progID, _, _ := readCString(data, off)

	d := Detection{Kind: KindCompObj, UserType: userType}
	if looksLikeProgID(progID) {
		d.ProgID = progID
	} else if looksLikeProgID(clipOrProg) {
		d.ProgID = clipOrProg
	} else if looksLikeProgID(userType) {
		d.ProgID = userType
	}
	if d.UserType == "" && d.ProgID == "" {
		return Detection{}, false
	}
	return d, true
}

func looksLikeProgID(s string) bool {
	if s == "" {
		return false
	}
	if strings.Contains(s, " ") {
		return false
	}
	return strings.Contains(s, ".")
}

func readCString(data []byte, offset int) (string, int, bool) {
	if offset < 0 || offset >= len(data) {
		return "", 0, false
	}
	end := offset
	for end < len(data) && data[end] != 0 {
		end++
	}
	if end >= len(data) {
		return "", 0, false
	}
	return string(data[offset:end]), end - offset + 1, true
}
