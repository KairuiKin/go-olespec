# API Freeze v1 (Draft)

## 1. 核心原则

- 公共 API 在 `v1.x` 内保持向后兼容。
- 基础库不管理 UI。
- 同时支持文件后端与纯内存后端。

## 2. Backend 抽象

```go
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
```

## 3. Open/Create API

```go
func Open(r storage.ReadBackend, opt OpenOptions) (*File, error)
func OpenReadWrite(rw storage.WriteBackend, opt OpenOptions) (*File, error)

func OpenFile(path string, opt OpenOptions) (*File, error)
func OpenBytes(buf []byte, opt OpenOptions) (*File, error)
func OpenBytesRW(buf []byte, opt OpenOptions) (*File, error)

func CreateInMemory(opt CreateOptions) (*File, error)
func CreateFile(path string, opt CreateOptions) (*File, error)
```

## 4. File/Tx API

```go
type File struct{}
type Tx struct{}

func (f *File) Begin(opt TxOptions) (*Tx, error)
func (f *File) Walk(fn func(Node) error) error
func (f *File) WalkEx(opt WalkOptions, fn func(WalkEvent) error) (WalkResult, error)
func (f *File) OpenStream(path string) (StreamReader, error)
func (f *File) SnapshotBytes() ([]byte, error) // only mem backend
func (f *File) Close() error

func (tx *Tx) PutStream(path string, r io.Reader, size int64) error
func (tx *Tx) Delete(path string) error
func (tx *Tx) Rename(oldPath, newPath string) error
func (tx *Tx) CreateStorage(path string) error
func (tx *Tx) Commit(ctx context.Context, opt CommitOptions) (*CommitResult, error)
func (tx *Tx) Revert() error
```

## 5. Node / Stream

```go
type Node struct {
    ID         NodeID
    Type       NodeType
    Path       string
    Name       string
    ParentID   NodeID
    Size       int64
    CLSID      [16]byte
    StateBits  uint32
    CreatedAt  int64
    ModifiedAt int64
    ChildCount int
}

type StreamReader interface {
    io.Reader
    io.ReaderAt
    io.Seeker
    io.Closer

    Size() int64
    Path() string
    NodeID() NodeID
}
```

## 6. Error Contract

- 所有对外错误必须是 `*OLEError` 或可 `errors.As` 到 `*OLEError`。
- 必须提供 `Code/Message/Path/Offset/Op/Cause`。
- `Offset` 不可用时固定 `-1`。

核心错误码：

- `INVALID_ARGUMENT`
- `NOT_FOUND`
- `CONFLICT`
- `READ_ONLY`
- `BAD_HEADER`
- `BAD_SECTOR`
- `BAD_FAT_CHAIN`
- `CYCLE_DETECTED`
- `OUT_OF_BOUNDS`
- `DIR_CORRUPT`
- `MINISTREAM_CORRUPT`
- `LIMIT_EXCEEDED`
- `DEPTH_EXCEEDED`
- `QUOTA_EXCEEDED`
- `TX_CLOSED`
- `COMMIT_FAILED`

## 7. Report Contract

- `Warnings/Repairs/FatalErr/Partial` 必须完整可序列化。
- `strict` 模式结构违规直接失败。
- `lenient` 模式尽力恢复并记录降级行为。

## 8. Deterministic Contract

- 同输入 + 同配置 + 同版本，`Walk`、`ExtractReport`、`Diagnostics` 输出顺序一致。

## 9. 兼容策略

- `v1` 内禁止破坏性字段移除或语义变更。
- 新字段仅追加，并给出默认值语义。
- 破坏性变更只允许在 `v2`。
