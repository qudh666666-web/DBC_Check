# 规则来源与边界

## 主基准

Vector Informatik GmbH：

- `Rules for Legacy Communication Descriptions`
- Technical Reference，CAN / LIN / FlexRay
- Version 1.12，Released，2023

工具主要实现该文档第3章DBC Attributes and Conventions及第6.1节System Signals。

## 项目培训覆盖

`can_rules.json`用于保存当前项目与通用Vector基线不同或更严格的准出要求，包括NM、UDS和XCP。

## 判定边界

- Vector文档说明其属于DBC属性总览；各MICROSAR BSW组件的详细语义仍以组件技术参考为准。
- E2E各Profile只做属性集合、范围和基础关系检查，不擅自补全Profile专属参数。
- 无CAN矩阵模式只能证明未命中现有结构/手册/项目规则，不能证明DBC与客户需求完全一致。
- Factor/Offset的客户精度需求仍需矩阵或协议作为依据。
- 实际能否成功转换为ARXML，最终仍应通过Vector转换工具验证。
