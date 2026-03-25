# OLECore Architecture (v1)

## 1. 目标

- 构建 Go 原生 OLE/CFB 基础库，支持磁盘与纯内存后端。
- 实现稳定的读取、写入、事务提交、递归提取与诊断报告。
- 以公开规范驱动实现：MS-CFB / MS-OLEPS / MS-OLEDS。

## 2. 范围

### v1 必做

- CFB 容器层完整读写（header/FAT/MiniFAT/dir/stream）。
- 属性集（OLEPS）解析与写回。
- 常见 OLE 对象载荷识别（Ole10Native/CompObj/Package）。
- 递归提取与去重、限流、容错报告。

### v1 非目标

- doc/xls/ppt 业务语义完整解析与高保真编辑。
- UI 组件与前端交互逻辑。

## 3. 分层

- `olecore`：文件能力层（无 UI 依赖）。
- `oledoc`：文档视图模型层（UI 无关 ViewModel）。
- `oleui-*`：终端适配层（web/winui/cli）。

依赖方向：`oleui-* -> oledoc -> olecore`

## 4. 目录建议

```text
/cmd/oletool
/pkg/olecfb
/pkg/oleps
/pkg/oleds
/pkg/olextract
/internal/cfbheader
/internal/cfbfat
/internal/cfbminifat
/internal/cfbdir
/internal/cfbstream
/internal/cfbtxn
/internal/storage
/internal/repair
/internal/bounds
/internal/hashing
```

## 5. 后端抽象

统一后端接口支持磁盘和内存：

- `ReadBackend`：`ReadAt/Size/Close`
- `WriteBackend`：`WriteAt/Truncate/Sync`

官方后端：

- `FileBackend`
- `MemBackend`（支持 `SnapshotBytes`）

## 6. 事务语义

- 同一 `File` 同时只允许一个活跃事务。
- `Begin -> Mutate -> Commit/Revert`。
- 提交失败不污染原状态。
- v1 提交策略默认 `FullRewrite`，并支持受限 `Incremental`：
  - 仅支持“单个已存在流、大小不变”的原位更新。
  - 不满足条件自动回退 `FullRewrite`。

## 7. Path 规范

- 根路径固定 `/`。
- 段编码采用 JSON Pointer 风格：`~1` 代表 `/`，`~0` 代表 `~`。
- 禁止空段和尾随 `/`（根除外）。
- 内部比较键：`UTF-8 NFC + casefold`。
- 写入时禁止“同父目录仅大小写不同”的重名。

## 8. Deterministic 规则

同输入 + 同配置 + 同版本，输出顺序必须稳定。

- 节点排序键：`(name_key, name_raw_utf8, dir_entry_id)`
- Artifact 排序键：`(depth, parent_path_key, kind_rank, path_key, sha256, source_node_id, local_index)`
- 诊断排序键：`(severity_rank, code, path_key, offset, op, local_index)`

## 9. 安全与健壮性

- 所有偏移、长度、计数先校验后分配。
- 默认启用配额：总字节、单对象大小、递归深度、对象数量。
- `strict/lenient` 双模式：
  - `strict`：结构违规即失败。
  - `lenient`：尽力恢复，并写入 `Warnings/Repairs`。
- fuzz 常态化，崩溃类缺陷阻断发布。

## 10. 质量门禁

- 规范矩阵覆盖（MUST/SHOULD -> 测试 ID）。
- `go test -race` 必过。
- 回归语料与差分测试必过。
- 性能回归阈值纳入 CI。
- 基准基线（当前）：`BenchmarkExtractFlat`、`BenchmarkExtractRecursive`、`BenchmarkCommitFullRewrite`、`BenchmarkCommitIncrementalInPlace`。
