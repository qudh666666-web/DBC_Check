# CAN矩阵-DBC一致性检查工具 v3.0.9

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
"%LocalAppData%\Programs\Python\Python312\python.exe" -m pip install -r requirements.txt
```

运行：

```bat
run_checker.bat
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

## v3.0.9 DBC命名、节点补全和安全写回

主界面新增“命名设置”和“节点补全”入口。

- 报文格式：`{原始报文名}_{标识}_0x{真实CAN ID}`，例如 `VehicleStatus_can1_0x101`。
- 信号格式：`{标识}_sig0x{所属报文真实CAN ID}{原始信号名}`，例如 `can1_sig0x101DataLenght`。
- ID只读展示并自动读取每条 `BO_` 的实际ID；扩展帧只保留实际CAN ID，不把 `0x80000000` 存储标志写入名称。
- 默认标识、报文覆盖、信号覆盖、仅报文/仅信号处理，以及多选对象批量设置均可在窗口中完成。
- JSON配置示例见 `rename_config.example.json`。应用后会保存原名到新名映射；重新打开生成后的DBC并加载同一配置，可继续把 `can1` 替换为 `can2`，不会叠加前缀。
- 写回默认另存为新DBC，预览绑定输入文件指纹；报文名检查全文件冲突，信号名检查所属报文内冲突。注释正文、枚举显示文字和未知语句不做全文替换。
- “节点补全”仅把 `BO_`发送者和 `SG_`接收者中明确出现、但未在 `BU_` 声明的真实节点补入；不把 `Vector__XXX` 或所有其他节点猜作接收者。

离线验证：

```bat
D:\Program Files\LibreOffice\program\python.exe test_dbc_transform.py
D:\Program Files\LibreOffice\program\python.exe self_test.py
```

图形界面仍需安装包含 Tcl/Tk 的 Python；若仅需运行解析和离线自测，缺少 tkinter 时程序会保留核心导入能力并明确拒绝启动 GUI。
