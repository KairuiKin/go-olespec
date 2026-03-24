package storage

type ReadBackend interface {
	ReadAt(p []byte, off int64) (n int, err error)
	Size() int64
	Close() error
}

type WriteBackend interface {
	ReadBackend
	WriteAt(p []byte, off int64) (n int, err error)
	Truncate(size int64) error
	Sync() error
}

type Kind string

const (
	KindFile Kind = "file"
	KindMem  Kind = "mem"
)

type Info struct {
	Kind     Kind
	ReadOnly bool
	Name     string
}

type Introspect interface {
	Info() Info
}
