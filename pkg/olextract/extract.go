package olextract

import "github.com/KairuiKin/go-olespec/pkg/olecfb"

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

