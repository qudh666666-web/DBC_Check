# Vector手册规则版说明（v3.0）

本版本以内置基线规则实现以下文档：

- **Rules for Legacy Communication Descriptions**
- Technical Reference，CAN / LIN / FlexRay
- Vector Informatik GmbH，Version **1.12**，Status **Released**

## 两层规则

1. **Vector手册基线**：直接内置在 `can_matrix_checker.py` 中，可在界面勾选/取消“启用Vector手册v1.12规则”。
2. **项目培训规则**：保存在 `can_rules.json`，继续执行项目规定的 NM、UDS、XCP 等覆盖规则。

项目规则优先用于项目准出判断。例如当前培训规范仍要求：

- NM：`GenMsgSendType=Cyclic`、周期大于0、`NmAsrMessage=Yes`。
- UDS接收请求：显式 `GenMsgCycleTime=0`、`GenMsgSendType=NotUsed`、`DiagConnection=57345`。
- XCP：项目规则禁止显式配置 `GenMsgCycleTime`。

## 已实现的Vector章节

- 3.1 General Attributes：BusType、VFrameFormat、波特率/采样点、属性对象范围。
- 3.2 COM：发送类型、周期、重复发送、Delay、StartValue、TimeoutTime_<Ecu>。
- 3.3 E2E：Profile、DataId、DataLength及基础映射一致性。
- 3.4 SecOC：split message必需属性、截断长度与字节边界。
- 3.5/3.6 NM：AUTOSAR NM、OSEK-NM、ID范围、计数、周期、Offset、CBV/SNI。
- 3.7/3.8 CanTp/DCM：TpTxIndex、诊断角色、DiagConnection、DiagFdOnly。
- 3.9 J1939：ProtocolType、扩展帧、动态DLC末端信号。
- 3.10 XCP/CDD：层属性与完整载荷Rx信号。
- 3.11 Update Bit：`<X>_UB`、1 bit、NoSigSendType、ILSupport。
- 3.12 Invalid Value：值表中的 `SNA`。
- 6.1 System Signal：同节点同名信号的长度和值表一致性。

## 范围占位处理

项目DBC经常将 `BA_DEF_` 数值范围写为 `0..0`，但实际 `BA_` 赋值符合Vector手册范围。本版处理方式：

- 对本地定义范围与Vector手册不一致：**每个属性只报一次警告**。
- 不再对每条报文/信号重复报“超出0..0”。
- 实际值超出Vector手册标准范围：仍报**错误**。
- 枚举定义名称/顺序与手册不同：统一报**警告**；数值槽位合法时不逐条误报。

## 边界

该Vector文档明确属于属性总览。E2E各Profile、SecOC算法和各MICROSAR BSW模块的细节仍应以对应组件技术参考及项目规范为准。
