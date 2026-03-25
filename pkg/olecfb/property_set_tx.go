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

func (tx *Tx) PutSummaryInformation(set *oleps.PropertySet) error {
	return tx.putNamedPropertySet(
		[]string{"/\u0005SummaryInformation", "/SummaryInformation"},
		oleps.FMTIDSummaryInformation,
		set,
	)
}

func (tx *Tx) PutDocumentSummaryInformation(set *oleps.PropertySet) error {
	return tx.putNamedPropertySet(
		[]string{"/\u0005DocumentSummaryInformation", "/DocumentSummaryInformation"},
		oleps.FMTIDDocumentSummaryInformation,
		set,
	)
}

func (tx *Tx) putNamedPropertySet(paths []string, formatID oleps.GUID, set *oleps.PropertySet) error {
	if tx == nil || tx.closed {
		return newError(ErrTxClosed, "transaction is closed", "property.put_named", "", -1, nil)
	}
	if set == nil {
		return newError(ErrInvalidArgument, "property set is nil", "property.put_named", "", -1, nil)
	}
	target := paths[0]
	for _, p := range paths {
		_, n, ok := tx.findNodeByPath(p)
		if ok {
			if !n.IsStream() {
				return newError(ErrConflict, "target path exists and is not a stream", "property.put_named", p, -1, nil)
			}
			target = p
			break
		}
	}

	stream := &oleps.Stream{ByteOrder: 0xFFFE}
	if id, node, ok := tx.findNodeByPath(target); ok && node.IsStream() {
		data, err := tx.loadStreamData(id, node.Size)
		if err != nil {
			return newError(ErrBadFATChain, "failed to read existing property set stream", "property.put_named", target, -1, err)
		}
		parsed, parseErr := oleps.Parse(data)
		if parseErr != nil {
			return newError(ErrDirCorrupt, "failed to parse existing property set stream", "property.put_named", target, -1, parseErr)
		}
		stream = parsed
	}
	ps := *set
	ps.FormatID = formatID
	replaced := false
	for i := range stream.Sets {
		if stream.Sets[i].FormatID == formatID {
			stream.Sets[i] = ps
			replaced = true
			break
		}
	}
	if !replaced {
		stream.Sets = append(stream.Sets, ps)
	}
	if stream.ByteOrder == 0 {
		stream.ByteOrder = 0xFFFE
	}
	return tx.PutPropertySet(target, stream)
}
