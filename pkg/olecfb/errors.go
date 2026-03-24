package olecfb

import (
	"errors"
	"fmt"
)

type ErrorCode string

const (
	ErrInvalidArgument   ErrorCode = "INVALID_ARGUMENT"
	ErrNotFound          ErrorCode = "NOT_FOUND"
	ErrConflict          ErrorCode = "CONFLICT"
	ErrReadOnly          ErrorCode = "READ_ONLY"
	ErrUnsupported       ErrorCode = "UNSUPPORTED"
	ErrBadHeader         ErrorCode = "BAD_HEADER"
	ErrBadSector         ErrorCode = "BAD_SECTOR"
	ErrBadFATChain       ErrorCode = "BAD_FAT_CHAIN"
	ErrCycleDetected     ErrorCode = "CYCLE_DETECTED"
	ErrOutOfBounds       ErrorCode = "OUT_OF_BOUNDS"
	ErrDirCorrupt        ErrorCode = "DIR_CORRUPT"
	ErrMiniStreamCorrupt ErrorCode = "MINISTREAM_CORRUPT"
	ErrLimitExceeded     ErrorCode = "LIMIT_EXCEEDED"
	ErrDepthExceeded     ErrorCode = "DEPTH_EXCEEDED"
	ErrQuotaExceeded     ErrorCode = "QUOTA_EXCEEDED"
	ErrTxClosed          ErrorCode = "TX_CLOSED"
	ErrCommitFailed      ErrorCode = "COMMIT_FAILED"
	ErrRevertFailed      ErrorCode = "REVERT_FAILED"
)

type OLEError struct {
	Code      ErrorCode
	Message   string
	Path      string
	Offset    int64
	Op        string
	Temporary bool
	Cause     error
}

func (e *OLEError) Error() string {
	if e == nil {
		return ""
	}
	base := fmt.Sprintf("%s: %s", e.Code, e.Message)
	if e.Path != "" {
		base += " path=" + e.Path
	}
	if e.Offset >= 0 {
		base += fmt.Sprintf(" off=%d", e.Offset)
	}
	if e.Op != "" {
		base += " op=" + e.Op
	}
	return base
}

func (e *OLEError) Unwrap() error { return e.Cause }

func IsCode(err error, code ErrorCode) bool {
	var oe *OLEError
	return errors.As(err, &oe) && oe.Code == code
}

func newError(code ErrorCode, msg, op, path string, offset int64, cause error) *OLEError {
	return &OLEError{
		Code:    code,
		Message: msg,
		Op:      op,
		Path:    path,
		Offset:  offset,
		Cause:   cause,
	}
}
