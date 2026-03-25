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

`olextract` 门面 API：

```go
func ExtractBackend(rb storage.ReadBackend, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error)
func ExtractBytes(buf []byte, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error)
func ExtractFile(path string, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error)
func ExtractReader(r io.Reader, openOpt olecfb.OpenOptions, extractOpt olecfb.ExtractOptions) (*olecfb.ExtractReport, error)
```

`ExtractBackend` 语义：

- 始终在返回前关闭 `ReadBackend`。
- `rb=nil` 返回 `INVALID_ARGUMENT`。

`oleds` 基础解析 API：

```go
type Ole10Native struct {
    FileName   string
    SourcePath string
    TempPath   string
    Payload    []byte
}

func Detect(streamPath string, data []byte) Detection
func ParseOle10Native(data []byte) (Ole10Native, bool)
```

`Open` 配额行为：

- `OpenOptions.MaxTotalBytes > 0` 时，若容器大小超限，`Open` 直接返回 `QUOTA_EXCEEDED`

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
func (f *File) OpenPropertySet(path string) (*oleps.Stream, error)
func (f *File) OpenSummaryInformation() (*oleps.PropertySet, error)
func (f *File) OpenDocumentSummaryInformation() (*oleps.PropertySet, error)

func (tx *Tx) PutStream(path string, r io.Reader, size int64) error
func (tx *Tx) Delete(path string) error
func (tx *Tx) Rename(oldPath, newPath string) error
func (tx *Tx) CreateStorage(path string) error
func (tx *Tx) PutPropertySet(path string, stream *oleps.Stream) error
func (tx *Tx) PutSummaryInformation(set *oleps.PropertySet) error
func (tx *Tx) PutDocumentSummaryInformation(set *oleps.PropertySet) error
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

`Artifact` 新增可选字段：

- `Raw []byte`: 当 `ExtractOptions.IncludeRaw=true` 时填充原始流字节；默认 `nil`
- `OLEFileName/OLESourcePath/OLETempPath`: 当 `DetectOLEDS` 或 `UnwrapOle10Native` 命中 `Ole10Native` 时填充来源元信息

`CommitResult` 字段：

- `BytesWritten`: 本次提交写入字节数
- `NewSize`: 提交后容器大小
- `BackendKind`: `mem` / `file`
- `StrategyUsed`: 实际使用的提交策略（`FullRewrite` 或 `Incremental`）

`Incremental`（v1）约束：

- 仅支持“单个已存在流、大小不变”的原位更新
- 其他场景自动回退 `FullRewrite`

`ExtractOptions` 追加字段：

- `UnwrapOle10Native bool`：开启后会解析 `Ole10Native` 流并把内嵌 payload 作为子 artifact 继续提取（含嵌套 OLE 递归）。

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
