# go-olespec

Core library for OLE/CFB specification-driven parsing, editing, and extraction in Go.

## Scope (current)

- `pkg/olecfb`: core contracts for CFB/OLE container operations
- `pkg/oledoc`: UI-agnostic document view model contracts
- `pkg/oleps`: OLE property set stream parser (minimal)
- `pkg/oleds`: placeholder for OLE object data support
- `pkg/olextract`: placeholder for recursive extraction support

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
- Property set parsing:
  - parse property set stream header and sets
  - read SummaryInformation common fields (string/int/time basic types)
