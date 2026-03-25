# go-olespec

Core library for OLE/CFB specification-driven parsing, editing, and extraction in Go.

## Scope (current)

- `pkg/olecfb`: core contracts for CFB/OLE container operations
- `pkg/oledoc`: UI-agnostic document view model contracts
- `pkg/oleps`: OLE property set stream parser (minimal)
- `pkg/oleds`: OLE object stream detector (Ole10Native/CompObj/Package)
- `pkg/olextract`: placeholder package (recursive extraction currently integrated in `pkg/olecfb`)

## Implemented so far (`pkg/olecfb`)

- CFB header parsing and validation (v3/v4)
- FAT loading from DIFAT header entries
- FAT chain traversal with cycle and bounds detection
- Directory stream parsing and node tree construction
- `Walk` / `WalkEx` traversal (DFS/BFS)
- Stream reading:
  - normal FAT streams
  - MiniFAT streams
- Strict/lenient parsing modes with warning report
- Basic extraction report with stream hashing (SHA-256)
- Recursive extraction for nested OLE streams with `ParentID/Children` graph
- OLE object detection in extraction (`DetectOLEDS`)
- Basic image signature detection in extraction (`DetectImages`)
- Optional raw payload embedding in artifacts (`IncludeRaw`)
- Property set parsing:
  - parse property set stream header and sets
  - read SummaryInformation common fields (string/int/time basic types)
- Transaction (v1):
  - `CreateStorage` / `PutStream` / `Delete` / `Rename` / `Commit` / `Revert`
  - `Commit` uses `FullRewrite` serializer and writes back to mem/file backend
  - `CommitOptions{Strategy: Incremental}` supports in-place updates for existing streams with unchanged size
  - other incremental cases transparently fallback to `FullRewrite`
