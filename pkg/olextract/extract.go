package olextract

import (
	"io"

	"github.com/KairuiKin/go-olespec/pkg/olecfb"
)

// ExtractBytes opens an OLE/CFB buffer and runs extraction with the provided options.
func ExtractBytes(buf []byte, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error) {
	f, err := olecfb.OpenBytes(buf, openOpt)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return f.Extract(extractOpt)
}

// ExtractFile opens an OLE/CFB file path and runs extraction with the provided options.
func ExtractFile(path string, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error) {
	f, err := olecfb.OpenFile(path, openOpt)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	return f.Extract(extractOpt)
}

// ExtractReader reads the entire input and runs extraction with the provided options.
func ExtractReader(r io.Reader, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error) {
	if r == nil {
		return nil, &olecfb.OLEError{
			Code:    olecfb.ErrInvalidArgument,
			Message: "reader is nil",
			Op:      "olextract.extract_reader",
			Offset:  -1,
		}
	}
	buf, err := io.ReadAll(r)
	if err != nil {
		return nil, &olecfb.OLEError{
			Code:    olecfb.ErrUnsupported,
			Message: "read input failed",
			Op:      "olextract.extract_reader",
			Offset:  -1,
			Cause:   err,
		}
	}
	return ExtractBytes(buf, openOpt, extractOpt)
}
