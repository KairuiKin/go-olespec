# OLEDoc Document Model v1

## 1. 目的

`oledoc` 是 `olecore` 与 UI 的稳定契约层，负责文档树、预览与增量变更协议，不包含具体 UI 代码。

## 2. DocumentModel

```go
type DocumentModel struct {
    Version     string
    DocumentID  string
    Source      SourceMeta
    RootID      string
    Nodes       map[string]DocNode
    Order       []string
    Diagnostics []Diagnostic
    Stats       ModelStats
}
```

## 3. 节点模型

```go
type DocNode struct {
    ID         string
    ParentID   string
    Path       string
    Name       string
    Type       NodeType
    Size       int64
    Depth      int
    HasChild   bool
    ChildCount int
    Tags       []string
    Meta       map[string]any
    State      NodeState
}
```

节点类型：`root/storage/stream/artifact`

节点状态：`ok/partial/broken/skipped`

## 4. 诊断模型

```go
type Diagnostic struct {
    Level   string
    Code    string
    Message string
    Path    string
    Offset  int64
    Op      string
}
```

要求：统一复用 `olecore` 错误码，不允许 UI 自定义主错误码。

## 5. 查询接口（只读）

```go
type QueryService interface {
    GetModel() (*DocumentModel, error)
    GetNode(id string) (*DocNode, error)
    GetChildren(parentID string, cursor string, limit int) (Page[DocNode], error)
    GetPreview(nodeID string, opt PreviewOptions) (*Preview, error)
    Find(q FindQuery) ([]FindHit, error)
    GetDiagnostics() ([]Diagnostic, error)
}
```

## 6. 变更协议（ChangeSet）

```go
type ChangeSet struct {
    TxID        string
    Applied     bool
    Revision    int64
    Mutations   []Mutation
    Diagnostics []Diagnostic
}

type Mutation struct {
    Kind      MutationKind
    NodeID    string
    ParentID  string
    Path      string
    OldPath   string
    Fields    []string
    Before    map[string]any
    After     map[string]any
}
```

Mutation 类型：`create/update/delete/move/rename`

## 7. 命令接口（事务化）

```go
type CommandService interface {
    Begin() (txID string, err error)
    CreateStorage(txID, parentPath, name string) (*ChangeSet, error)
    PutStream(txID, path string, r io.Reader, size int64) (*ChangeSet, error)
    Delete(txID, path string) (*ChangeSet, error)
    Rename(txID, oldPath, newPath string) (*ChangeSet, error)
    Commit(txID string) (*ChangeSet, error)
    Revert(txID string) (*ChangeSet, error)
}
```

## 8. 强约束

- 同输入/同配置下，`DocumentModel.Order` 必须 deterministic。
- `Meta` 只允许轻量字段，不允许大二进制。
- 所有命令必须返回 `ChangeSet`，UI 禁止依赖“全量重刷”。
