package olecfb

import (
	"errors"
	"io"

	"github.com/KairuiKin/go-olespec/pkg/oleps"
)

func (f *File) OpenPropertySet(path string) (*oleps.Stream, error) {
	if f == nil {
		return nil, newError(ErrInvalidArgument, "file is nil", "property.open", path, -1, nil)
	}
	sr, err := f.OpenStream(path)
	if err != nil {
		return nil, err
	}
	defer sr.Close()
	data, err := io.ReadAll(sr)
	if err != nil {
		return nil, newError(ErrBadFATChain, "failed to read property set stream", "property.open", path, -1, err)
	}
	stream, err := oleps.Parse(data)
	if err != nil {
		return nil, newError(ErrDirCorrupt, "failed to parse property set stream", "property.open", path, -1, err)
	}
	return stream, nil
}

func (f *File) OpenSummaryInformation() (*oleps.PropertySet, error) {
	if f == nil {
		return nil, newError(ErrInvalidArgument, "file is nil", "property.summary", "", -1, nil)
	}
	paths := []string{"/\u0005SummaryInformation", "/SummaryInformation"}
	var lastErr error
	for _, p := range paths {
		pss, err := f.OpenPropertySet(p)
		if err != nil {
			lastErr = err
			continue
		}
		if set, ok := pss.SummaryInformation(); ok {
			return set, nil
		}
		lastErr = newError(ErrNotFound, "summary information set not found in stream", "property.summary", p, -1, nil)
	}
	if lastErr == nil {
		lastErr = errors.New("summary information stream not found")
	}
	if oe, ok := lastErr.(*OLEError); ok {
		return nil, oe
	}
	return nil, newError(ErrNotFound, "summary information not found", "property.summary", "", -1, lastErr)
}
