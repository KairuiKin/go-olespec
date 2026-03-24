package olecfb

import (
	"bytes"
	"fmt"
	"io"
)

func readNormalStreamData(
	readAt func([]byte, int64) (int, error),
	fileSize int64,
	fat []uint32,
	hdr *cfbHeader,
	startSector uint32,
	size int64,
	maxChainLength int,
) ([]byte, error) {
	if hdr == nil {
		return nil, newError(ErrBadHeader, "header is nil", "stream.read", "", -1, nil)
	}
	if size == 0 {
		return nil, nil
	}
	chain, err := walkFATChain(fat, startSector, maxChainLength)
	if err != nil {
		return nil, newError(ErrBadFATChain, "failed to resolve stream chain", "stream.read", "", -1, err)
	}
	if len(chain) == 0 {
		return nil, newError(ErrBadFATChain, "stream chain is empty", "stream.read", "", -1, nil)
	}

	sectorSize := int64(1 << hdr.SectorShift)
	expectedSectors := int((size + sectorSize - 1) / sectorSize)
	if len(chain) < expectedSectors {
		return nil, newError(ErrBadFATChain, fmt.Sprintf("stream chain too short: have %d need %d", len(chain), expectedSectors), "stream.read", "", -1, nil)
	}

	out := make([]byte, size)
	written := int64(0)
	for _, sid := range chain {
		if written >= size {
			break
		}
		off := sectorOffset(sid, sectorSize)
		if off < 0 || off+sectorSize > fileSize {
			return nil, newError(ErrOutOfBounds, "stream sector out of bounds", "stream.read", "", off, nil)
		}
		toRead := sectorSize
		if rem := size - written; rem < toRead {
			toRead = rem
		}
		if err := readFullAt(readAt, out[written:written+toRead], off); err != nil {
			return nil, newError(ErrBadFATChain, "failed to read stream sector", "stream.read", "", off, err)
		}
		written += toRead
	}
	return out, nil
}

type bytesStreamReader struct {
	nodeID NodeID
	path   string
	size   int64
	reader *bytes.Reader
}

func newBytesStreamReader(n Node, data []byte) *bytesStreamReader {
	return &bytesStreamReader{
		nodeID: n.ID,
		path:   n.Path,
		size:   int64(len(data)),
		reader: bytes.NewReader(data),
	}
}

func (r *bytesStreamReader) Read(p []byte) (int, error) { return r.reader.Read(p) }
func (r *bytesStreamReader) ReadAt(p []byte, off int64) (int, error) {
	return r.reader.ReadAt(p, off)
}
func (r *bytesStreamReader) Seek(offset int64, whence int) (int64, error) {
	return r.reader.Seek(offset, whence)
}
func (r *bytesStreamReader) Close() error { return nil }
func (r *bytesStreamReader) Size() int64  { return r.size }
func (r *bytesStreamReader) Path() string { return r.path }
func (r *bytesStreamReader) NodeID() NodeID {
	return r.nodeID
}

var _ StreamReader = (*bytesStreamReader)(nil)
var _ io.Reader = (*bytesStreamReader)(nil)
