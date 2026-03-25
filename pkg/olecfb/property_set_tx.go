package olecfb

import (
	"bytes"

	"github.com/KairuiKin/go-olespec/pkg/oleps"
)

func (tx *Tx) PutPropertySet(path string, stream *oleps.Stream) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "property.put", path, -1, nil)
	}
	if stream == nil {
		return newError(ErrInvalidArgument, "property set stream is nil", "property.put", path, -1, nil)
	}
	buf, err := oleps.Marshal(stream)
	if err != nil {
		return newError(ErrInvalidArgument, "failed to marshal property set stream", "property.put", path, -1, err)
	}
	return tx.PutStream(path, bytes.NewReader(buf), int64(len(buf)))
}

