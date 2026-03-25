# Development Roadmap

## Vision

Build a production-grade, spec-driven Go OLE core library that supports reliable extraction, mutation, and future doc/xls/ppt editing foundations.

## Milestone M1: v1.1 Hardening (Current -> 4 weeks)

- Scope:
  - Extraction stability: recursive Ole10Native/OLE recursion/limits fuzzing.
  - Write-out stability: flat/tree/manifest path safety and cross-platform path normalization.
  - API consistency: finalize `olextract` facade and error contract.
- Deliverables:
  - Fuzz targets for `olecfb.Extract`, `oleds.ParseOle10Native`, `WriteArtifacts`.
  - Golden corpus + differential regression suite.
  - CI benchmark thresholds for `Extract` and `Commit`.
- Exit criteria:
  - `go test ./...` + fuzz smoke pass in CI.
  - No known high/critical path traversal or corruption issues.

## Milestone M2: v1.2 Interop & Coverage (4 -> 8 weeks)

- Scope:
  - Broader OLEDS parsing (CompObj structure fields, Package metadata).
  - Property-set interoperability corpus (Office/WPS variants).
  - Deterministic output verification across platforms.
- Deliverables:
  - Extended `oleds` parse APIs and compatibility tests.
  - Corpus replay tool under `cmd/` for batch validation. (initial CLI + baseline diff/gate + error-code regression gates 已落地：`cmd/corpus-replay`)
  - Structured extraction manifest schema freeze.
- Exit criteria:
  - Interop pass rate target >= 95% on curated corpus.

## Milestone M3: v2 Preparation (8 -> 16 weeks)

- Scope:
  - Incremental commit strategy expansion (beyond single-stream fixed-size update).
  - Snapshot/rollback primitives for safer multi-step edits.
  - Large-file streaming extraction mode (reduce peak memory).
- Deliverables:
  - `Incremental` capability matrix and staged enable flags.
  - Transaction reliability tests for mixed edit workloads.
  - Streaming extraction API proposal + prototype.
- Exit criteria:
  - Clear v2 API proposal with migration notes from v1.

## Engineering Tracks (Parallel)

- Security:
  - Path/manifest safety, malformed input hardening, fuzz first policy.
- Performance:
  - Bench trend tracking; avoid >10% regressions on core benchmarks.
- Documentation:
  - Keep `api-freeze-v1.md` and `spec-matrix.yaml` synchronized with every feature.
- Release:
  - Monthly patch train; semantic versioning discipline; changelog automation.
