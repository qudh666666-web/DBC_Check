# v3.0 规则配置说明

工具采用两层规则：

```text
内置Vector Technical Reference v1.12基线
+ can_rules.json项目培训覆盖规则
```

`can_rules.json`只需要维护项目差异，不需要重复录入Vector手册中的标准属性范围。

## 当前项目覆盖

### NM

```text
GenMsgSendType = Cyclic
GenMsgCycleTime > 0
NmAsrMessage / NmMessage = Yes
```

### UDS接收请求

```text
GenMsgCycleTime = 0（显式）
GenMsgSendType = NotUsed（显式）
DiagConnection = 57345（显式）
```

### XCP

```text
禁止显式配置 GenMsgCycleTime
```

Vector手册本身还会检查XCP/CCP的ILSupport、NM、诊断层属性以及完整载荷Rx信号。

## 公共字段

```json
{
  "id": "RULE_ID",
  "enabled": true,
  "severity": "错误",
  "target": "DBC",
  "type": "规则类型",
  "field_name": "报告中的差异项",
  "description": "规则说明"
}
```

`severity`支持：`错误`、`警告`、`提示`。

## 其他可配置规则类型

- `signal_name_length_allowed`
- `signal_name_length_range`
- `message_name_require_signal_patterns`
- `signal_binary_enum_length`
- `message_cycle_allowed`
- `message_name_dlc_allowed`
- `message_dlc_max`
- `message_name_frame_format_allowed`
- `message_name_id_range`
- `message_name_send_type_forbidden`

周期白名单只有项目明确给出时再启用。
