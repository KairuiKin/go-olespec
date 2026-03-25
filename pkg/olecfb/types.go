package olecfb

import (
	"io"
	"time"
)

type NodeType uint8

const (
	NodeUnknown NodeType = iota
	NodeStorage
	NodeStream
	NodeRoot
)

type NodeID uint32

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

func (n Node) IsStream() bool  { return n.Type == NodeStream }
func (n Node) IsStorage() bool { return n.Type == NodeStorage || n.Type == NodeRoot }

type ParseMode int

const (
	Strict ParseMode = iota
	Lenient
)

type CommitStrategy int

const (
	FullRewrite CommitStrategy = iota
	Incremental
)

type OpenOptions struct {
	Mode           ParseMode
	MaxObjectCount int
	MaxTotalBytes  int64
	MaxStreamBytes int64
	MaxChainLength int
	MaxRecursion   int
}

type CreateOptions struct {
	Deterministic bool
}

type TxOptions struct {
	Deterministic bool
}

type CommitOptions struct {
	Strategy CommitStrategy
	Sync     bool
}

type CommitResult struct {
	BytesWritten int64
	NewSize      int64
	BackendKind  string
	StrategyUsed CommitStrategy
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

type StreamWriter interface {
	io.Writer
	io.WriterAt
	io.Seeker
	io.Closer

	Size() int64
	Path() string
	NodeID() NodeID
}

type Severity string

const (
	SeverityInfo    Severity = "info"
	SeverityWarning Severity = "warning"
	SeverityError   Severity = "error"
)

type Warning struct {
	Code     ErrorCode
	Message  string
	Path     string
	Offset   int64
	Op       string
	Severity Severity
}

type RepairAction string

const (
	RepairSkippedNode       RepairAction = "skipped_node"
	RepairClampedSize       RepairAction = "clamped_size"
	RepairBrokeFATCycle     RepairAction = "broke_fat_cycle"
	RepairTrimmedOutOfBound RepairAction = "trimmed_out_of_bound"
	RepairFallbackMiniFAT   RepairAction = "fallback_minifat"
)

type RepairRecord struct {
	Action  RepairAction
	Code    ErrorCode
	Path    string
	Offset  int64
	Before  string
	After   string
	Message string
}

type ParseStats struct {
	Duration        time.Duration
	NodesTotal      int
	StreamsTotal    int
	StoragesTotal   int
	BytesRead       int64
	BytesLogical    int64
	MaxDepthReached int
}

type Report struct {
	Mode     ParseMode
	Stats    ParseStats
	Warnings []Warning
	Repairs  []RepairRecord
	FatalErr *OLEError
	Partial  bool
}

type WalkOptions struct {
	MaxDepth    int
	IncludeRoot bool
	Order       WalkOrder
}

type WalkOrder string

const (
	WalkDFS WalkOrder = "dfs"
	WalkBFS WalkOrder = "bfs"
)

type WalkEvent struct {
	Node  Node
	Depth int
	Index int
}

type WalkResult struct {
	Visited  int
	Skipped  int
	MaxDepth int
	Duration time.Duration
	Warnings []Warning
}

type ArtifactKind string

const (
	ArtifactOLEFile ArtifactKind = "ole_file"
	ArtifactStream  ArtifactKind = "stream"
	ArtifactImage   ArtifactKind = "image"
	ArtifactOleObj  ArtifactKind = "ole_object"
	ArtifactUnknown ArtifactKind = "unknown"
)

type ArtifactStatus string

const (
	ArtifactOK      ArtifactStatus = "ok"
	ArtifactPartial ArtifactStatus = "partial"
	ArtifactFailed  ArtifactStatus = "failed"
	ArtifactSkipped ArtifactStatus = "skipped"
)

type Artifact struct {
	ID            string
	Kind          ArtifactKind
	Status        ArtifactStatus
	Path          string
	MediaType     string
	Size          int64
	SHA256        string
	Raw           []byte
	Depth         int
	ParentID      string
	Children      int
	SourceNodeID  NodeID
	OLEFileName   string
	OLESourcePath string
	OLETempPath   string
	Error         *OLEError
	Note          string
}

type ExtractLimits struct {
	MaxDepth        int
	MaxArtifacts    int
	MaxTotalBytes   int64
	MaxArtifactSize int64
}

type ExtractOptions struct {
	Mode              ParseMode
	Limits            ExtractLimits
	IncludeRaw        bool
	DetectImages      bool
	DetectOLEDS       bool
	UnwrapOle10Native bool
	Deduplicate       bool
}

type ExtractStats struct {
	Duration         time.Duration
	ArtifactsTotal   int
	ArtifactsOK      int
	ArtifactsPartial int
	ArtifactsFailed  int
	DedupHits        int
	BytesExported    int64
	MaxDepthReached  int
}

type ExtractReport struct {
	Stats     ExtractStats
	Artifacts []Artifact
	Warnings  []Warning
	Repairs   []RepairRecord
	Partial   bool
	FatalErr  *OLEError
}
