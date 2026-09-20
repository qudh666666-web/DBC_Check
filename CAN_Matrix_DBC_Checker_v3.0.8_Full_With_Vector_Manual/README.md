# CAN矩阵-DBC一致性检查工具 v3.0.7

Python离线图形工具，支持：

- CAN通信矩阵（xlsx/xlsm/csv）与DBC自动对比。
- 无CAN矩阵时，仅检查DBC结构、Vector手册规则和项目培训规则。
- 导出Excel/CSV报告。
- 检查界面、完整详情和导出报告同时显示CAN ID的十六进制值与十进制值，便于搜索DBC代码。
- 可拖动调整结果表与完整详情框大小。
- 长内容自动换行，单击结果可查看完整详情。

## 运行环境

Python 3.9或更高版本。Excel功能需要：

```bat
pip install openpyxl
```

运行：

```bat
python can_matrix_checker.py
```

或双击 `run_checker.bat`。

## 规则层级

### 1. Vector手册内置基线

基准为Vector Technical Reference **Rules for Legacy Communication Descriptions, Version 1.12**。
界面默认勾选“启用Vector手册v1.12规则”。

覆盖General、COM、E2E、SecOC、AUTOSAR/OSEK NM、CanTp/DCM、J1939、XCP、Update Bit、Invalid Value和System Signal。

### 2. 项目培训覆盖规则

`can_rules.json`保留项目特殊要求：

- NM：Cyclic、周期大于0、NmAsrMessage=Yes。
- UDS接收请求：CycleTime=0、SendType=NotUsed、DiagConnection=57345，且要求显式赋值。
- XCP：禁止显式配置GenMsgCycleTime。

## BA_DEF_ 0..0占位范围

对Vector已定义标准范围的属性：

- 本地BA_DEF_范围不一致：每个属性只报一次警告。
- 实际BA_值：按Vector手册范围检查。
- 因此GenMsgDelayTime=10、GenMsgCycleTime=500等不会再因为模板定义0..0而重复产生大量警告。

## 测试样例

- `sample_vector_manual_bad.dbc`：故意触发Vector手册规则。
- `sample_project_rules_bad.dbc`：故意触发项目培训规则。
- `sample_matrix.csv` + `sample.dbc`：基础矩阵对比样例。

详见 `VECTOR_MANUAL_RULES.md` 和 `RULES_GUIDE.md`。


## v3.0.3 E2EDataLength 智能建议
- 若报文只有一个 `SIG_GROUP_`，优先按该 SignalGroup 成员在 PDU 中覆盖的字节区间估算 `E2EDataLength`。
- 例如保护组序列化长度为 2 Byte，则候选值为 `16 bit`。
- 若存在多个 SignalGroup，会分别列出候选值，不擅自选择。
- 无法识别保护组时才退回 `DLC × 8`，并标记为低置信度兜底。
- 建议用生成代码 `transformationBuffer[N]` 复核：重点检查 `N × 8 bit`。


## v3.0.7 一键修改E2E
- 检查完成后，如果存在可唯一确定的 `E2EDataLength` 高置信度候选，顶部会启用“一键修改E2E”按钮。
- 目前只自动修改“报文只有一个可解析 SignalGroup”的情况。
- 点击后会先在DBC同目录创建 `*.before_E2E_fix_时间戳.dbc` 备份，再直接写回当前DBC。
- 修改完成后程序会自动重新检查。
- 多个 SignalGroup、仅能用 `DLC × 8` 推测等低置信度情况不会自动写回，仍保留人工确认。


## v3.0.7 E2E单条/批量修改
- 结果表“操作”列中的“一键修改”：只修改当前报文的高置信度 E2EDataLength。
- 顶部“全部一键修改E2E”：一次修改全部可安全确定的项目。
- 每次写回前自动备份 DBC；低置信度候选仅提示，不自动修改。


## v3.0.7：客户E2E ID表导入

界面新增“E2E ID表”输入。支持 Excel/CSV，并自动识别类似：

```text
TX | DATA ID | Rx | DATAID
BcpSysStsSigGrp | 0x062D | LonAccrSigGrp | 0x0103
```

执行“开始检查”后，会按 SignalGroup 名称定位DBC，并核对 E2EDataId。主要规则：

- 客户表同一SignalGroup出现不同Data ID：错误
- 客户表SignalGroup在DBC中不存在：错误
- DBC中对应E2EDataId缺失：错误，并显示客户建议值
- DBC Data ID与客户表不一致：错误
- 客户表Data ID为“-”/空白：提示，跳过数值核对

可点击“查看E2E映射”确认程序识别到的TX/Rx和DATA ID列。

## v3.0.8 E2E ID表多路/多工作表兼容

- 同一Sheet可识别多组 TX/RX + DATA ID 列。
- 一个Excel里有多路E2E工作表时，可选择“全部可识别工作表”统一导入。
- 客户表允许不完整：未列出的DBC E2E组不判错；有SignalGroup但Data ID为空/“-”只提示。
- 同一SignalGroup在不同路/不同Sheet给出不同Data ID时，报冲突并指出来源。
