# go-olespec

Core library for OLE/CFB specification-driven parsing, editing, and extraction in Go.

## Scope (current)

- `pkg/olecfb`: core contracts for CFB/OLE container operations
- `pkg/oledoc`: UI-agnostic document view model contracts
- `pkg/oleps`: OLE property set stream parser (minimal)
- `pkg/oleds`: OLE object stream detection + Ole10Native parser
- `pkg/olextract`: extraction-oriented convenience facade (`ExtractBackend`/`ExtractBytes`/`ExtractFile`/`ExtractReader`) and artifact write-out helpers (`WriteArtifacts`/`Extract*ToDir`, flat/tree layout + manifest + optional atomic publish)
- `cmd/corpus-replay`: batch corpus replay CLI for extraction pass/fail and coverage statistics

## Corpus Replay CLI

```bash
go run ./cmd/corpus-replay -root ./samples -ext .doc,.xls,.ppt,.cfb -mode lenient -output ./report.json
```

```bash
go run ./cmd/corpus-replay -root ./samples -ext .cfb -baseline ./baseline.json -max-newly-failed 0 -min-pass-rate 0.98
```

```bash
go run ./cmd/corpus-replay -root ./samples -ext .cfb -baseline ./baseline.json -max-new-error-codes 0 -max-error-code-regressions 0 -deny-error-codes BAD_HEADER,DIR_CORRUPT
```

```bash
go run ./cmd/corpus-replay -root ./samples -ext .cfb -run-id $(git rev-parse --short HEAD) -trend-dir ./reports/history -trend-limit 30
```

```bash
go run ./cmd/corpus-replay -root ./samples -ext .cfb -trend-dir ./reports/history -max-pass-rate-drop 0.02 -max-failed-increase 0
```

```bash
go run ./cmd/corpus-replay -root ./samples -ext .cfb -trend-dir ./reports/history -run-id $(git rev-parse --short HEAD) -save-trend
```

```bash
go run ./cmd/corpus-replay -root ./samples -ext .cfb -baseline ./baseline.json -max-new-files 0 -max-removed-files 0 -max-newly-partial 0
```

## Implemented so far (`pkg/olecfb`)

- CFB header parsing and validation (v3/v4)
- FAT loading from DIFAT header entries
- FAT chain traversal with cycle and bounds detection
- Directory stream parsing and node tree construction
- `Walk` / `WalkEx` traversal (DFS/BFS)
- Stream reading:
  - normal FAT streams
  - MiniFAT streams
  - enforce `OpenOptions.MaxStreamBytes` on stream open/read
- Strict/lenient parsing modes with warning report
- enforce `OpenOptions.MaxTotalBytes` at open-time
- Basic extraction report with stream hashing (SHA-256)
- Recursive extraction for nested OLE streams with `ParentID/Children` graph
- OLE object detection in extraction (`DetectOLEDS`)
- Optional Ole10Native payload unwrapping and recursive extraction (`UnwrapOle10Native`)
- Basic image signature detection in extraction (`DetectImages`)
- Optional raw payload embedding in artifacts (`IncludeRaw`)
- Structured Ole10Native metadata on artifacts (`OLEFileName/OLESourcePath/OLETempPath`)
- Extract limit defaults can inherit from `OpenOptions` quotas
- Property set parsing:
  - parse property set stream header and sets
  - read SummaryInformation / DocumentSummaryInformation fields (string/int/time basic types)
  - property mutation helpers (`SetString/SetInt64/SetBool/SetTime/Delete`)
  - marshal/write property set streams via transaction (`Tx.PutPropertySet`)
  - convenience write helpers: `Tx.PutSummaryInformation` / `Tx.PutDocumentSummaryInformation`
- Transaction (v1):
  - `CreateStorage` / `PutStream` / `Delete` / `Rename` / `Commit` / `Revert`
  - `Commit` uses `FullRewrite` serializer and writes back to mem/file backend
  - `CommitOptions{Strategy: Incremental}` supports in-place updates for a single existing stream with unchanged size
  - other incremental cases transparently fallback to `FullRewrite`
