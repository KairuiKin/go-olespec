package oledoc

import (
	"io"
	"time"
)

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

type SourceMeta struct {
	BackendKind string
	Name        string
	Size        int64
	ReadOnly    bool
	ParseMode   string
}

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

type NodeType string

const (
	NodeRoot     NodeType = "root"
	NodeStorage  NodeType = "storage"
	NodeStream   NodeType = "stream"
	NodeArtifact NodeType = "artifact"
)

type NodeState string

const (
	StateOK      NodeState = "ok"
	StatePartial NodeState = "partial"
	StateBroken  NodeState = "broken"
	StateSkipped NodeState = "skipped"
)

type Diagnostic struct {
	Level   string
	Code    string
	Message string
	Path    string
	Offset  int64
	Op      string
}

type ModelStats struct {
	NodeTotal     int
	StreamTotal   int
	StorageTotal  int
	ArtifactTotal int
	WarningTotal  int
	ErrorTotal    int
	MaxDepth      int
}

type Page[T any] struct {
	Items      []T
	NextCursor string
	HasMore    bool
}

type PreviewOptions struct {
	MaxBytes int
	Timeout  time.Duration
}

type Preview struct {
	Kind      string
	Text      string
	Binary    []byte
	MediaType string
	Meta      map[string]any
}

type FindQuery struct {
	PathPrefix string
	NameLike   string
	Tag        string
	Limit      int
}

type FindHit struct {
	NodeID string
	Path   string
	Name   string
	Score  float64
}

type QueryService interface {
	GetModel() (*DocumentModel, error)
	GetNode(id string) (*DocNode, error)
	GetChildren(parentID string, cursor string, limit int) (Page[DocNode], error)
	GetPreview(nodeID string, opt PreviewOptions) (*Preview, error)
	Find(q FindQuery) ([]FindHit, error)
	GetDiagnostics() ([]Diagnostic, error)
}

type ChangeSet struct {
	TxID        string
	Applied     bool
	Revision    int64
	Mutations   []Mutation
	Diagnostics []Diagnostic
}

type Mutation struct {
	Kind     MutationKind
	NodeID   string
	ParentID string
	Path     string
	OldPath  string
	Fields   []string
	Before   map[string]any
	After    map[string]any
}

type MutationKind string

const (
	MutCreate MutationKind = "create"
	MutUpdate MutationKind = "update"
	MutDelete MutationKind = "delete"
	MutMove   MutationKind = "move"
	MutRename MutationKind = "rename"
)

type CommandService interface {
	Begin() (txID string, err error)
	CreateStorage(txID, parentPath, name string) (*ChangeSet, error)
	PutStream(txID, path string, r io.Reader, size int64) (*ChangeSet, error)
	Delete(txID, path string) (*ChangeSet, error)
	Rename(txID, oldPath, newPath string) (*ChangeSet, error)
	Commit(txID string) (*ChangeSet, error)
	Revert(txID string) (*ChangeSet, error)
}
