"""受控 DBC 命名、节点补全和保留原文写回。

这个模块只修改能够按 DBC 语句结构定位的名称/节点，不重建整个 Database，
因此未知语句、注释、值表显示文字和换行布局可以原样保留。
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


IDENTIFIER_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
BO_RE = re.compile(r"^(?P<prefix>\s*BO_\s+)(?P<raw_id>\d+)(?P<middle>\s+)(?P<name>[^:]+?)(?P<suffix>\s*:\s*.*)$")
SG_RE = re.compile(r"^(?P<prefix>\s*SG_\s+)(?P<name>\S+)(?P<suffix>\s+.*)$")
ATTRIBUTE_DEFINITION_RE = re.compile(
    r'^\s*BA_DEF_\s+(?:(?P<scope>BU_|BO_|SG_|EV_)\s+)?"(?P<name>[^"]+)"\s+'
    r'(?:INT|HEX|FLOAT|STRING|ENUM)\b.*;\s*$'
)
ATTRIBUTE_ASSIGNMENT_RE = re.compile(
    r'^\s*BA_\s+"(?P<name>[^"]+)"\s+(?:(?P<scope>BU_|BO_|SG_|EV_)\s+)?(?P<rest>.+?)\s*;\s*$'
)


def format_can_id(can_id: int) -> str:
    return f"0x{can_id:X}"


def normalize_raw_id(raw_id: int) -> Tuple[int, str]:
    can_id = raw_id & 0x1FFFFFFF if raw_id & 0x80000000 else raw_id
    frame = "extended" if raw_id & 0x80000000 or can_id > 0x7FF else "standard"
    return can_id, frame


def object_key(raw_id: int) -> str:
    can_id, frame = normalize_raw_id(raw_id)
    return f"{frame}:0x{can_id:X}"


def signal_key(message_key: str, signal_name: str) -> str:
    return f"{message_key}|{signal_name}"


def file_fingerprint(path: str) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def validate_identifier(value: str) -> str:
    value = str(value or "").strip()
    if not IDENTIFIER_RE.fullmatch(value):
        raise ValueError(f"命名标识“{value}”不合法；只能使用字母、数字、下划线，且不能以数字开头。")
    return value


def default_rename_config() -> Dict[str, Any]:
    return {
        "version": 1,
        "default_identifier": "can1",
        "targets": {"messages": True, "signals": True},
        "message_overrides": {},
        "signal_overrides": {},
        "mapping": {"source_fingerprint": "", "target_fingerprint": "", "messages": {}, "signals": {}},
    }


def load_rename_config(path: Optional[str] = None) -> Dict[str, Any]:
    config = default_rename_config()
    if path and Path(path).is_file():
        loaded = json.loads(Path(path).read_text(encoding="utf-8-sig"))
        if not isinstance(loaded, dict):
            raise ValueError("命名配置必须是 JSON 对象。")
        config.update(loaded)
    config.setdefault("targets", {})
    config["targets"] = {"messages": True, "signals": True, **config["targets"]}
    config.setdefault("message_overrides", {})
    config.setdefault("signal_overrides", {})
    config.setdefault("mapping", {"source_fingerprint": "", "target_fingerprint": "", "messages": {}, "signals": {}})
    config["mapping"] = {
        "source_fingerprint": "",
        "target_fingerprint": "",
        "messages": {},
        "signals": {},
        **config.get("mapping", {}),
    }
    validate_identifier(config.get("default_identifier", ""))
    for key, value in config["message_overrides"].items():
        config["message_overrides"][key] = validate_identifier(value)
    for key, value in config["signal_overrides"].items():
        config["signal_overrides"][key] = validate_identifier(value)
    return config


def save_rename_config(path: str, config: Dict[str, Any]) -> None:
    config = dict(config)
    config["default_identifier"] = validate_identifier(config.get("default_identifier", ""))
    config["version"] = 1
    Path(path).write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass(frozen=True)
class MessageRecord:
    raw_id: int
    can_id: int
    frame_format: str
    key: str
    name: str
    line_number: int


@dataclass(frozen=True)
class SignalRecord:
    message_key: str
    raw_id: int
    can_id: int
    frame_format: str
    name: str
    line_number: int


@dataclass
class RenamePreviewItem:
    object_type: str
    object_key: str
    raw_id: int
    can_id: int
    frame_format: str
    old_name: str
    new_name: str
    identifier: str
    status: str
    reason: str = ""


@dataclass
class RenamePlan:
    source_path: str
    source_fingerprint: str
    items: List[RenamePreviewItem]
    message_records: Tuple[MessageRecord, ...] = field(default_factory=tuple)
    signal_records: Tuple[SignalRecord, ...] = field(default_factory=tuple)

    @property
    def executable_items(self) -> List[RenamePreviewItem]:
        return [item for item in self.items if item.status == "可执行"]

    @property
    def blocked_items(self) -> List[RenamePreviewItem]:
        return [item for item in self.items if item.status == "阻塞"]


@dataclass(frozen=True)
class DbcRepairItem:
    """一条可审阅的 DBC 属性修复建议。

    未定义属性不能靠名称猜测类型或作用域。默认仅报告为“待用户配置”；
    用户明确选择删除时，才会把对应 BA_ 语句从另存副本中移除。
    """

    rule_id: str
    line_number: int
    attribute_name: str
    scope: str
    raw_line: str
    status: str
    reason: str


@dataclass
class DbcRepairPlan:
    source_path: str
    source_fingerprint: str
    items: List[DbcRepairItem]

    @property
    def blocked_items(self) -> List[DbcRepairItem]:
        return [item for item in self.items if item.status != "可执行"]


def _read_text(path: str) -> Tuple[str, str]:
    raw = Path(path).read_bytes()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig"), "utf-8-sig"
    for encoding in ("utf-8", "gb18030", "latin-1"):
        try:
            return raw.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    raise ValueError("DBC 文本编码无法识别，无法修改。")


def _line_without_eol(line: str) -> str:
    return line.rstrip("\r\n")


def _records(lines: Sequence[str]) -> Tuple[List[MessageRecord], List[SignalRecord]]:
    messages: List[MessageRecord] = []
    signals: List[SignalRecord] = []
    current: Optional[MessageRecord] = None
    for line_number, raw_line in enumerate(lines, start=1):
        line = _line_without_eol(raw_line)
        match = BO_RE.match(line)
        if match:
            raw_id = int(match.group("raw_id"))
            can_id, frame = normalize_raw_id(raw_id)
            current = MessageRecord(raw_id, can_id, frame, object_key(raw_id), match.group("name").strip(), line_number)
            messages.append(current)
            continue
        if current is not None:
            match = SG_RE.match(line)
            if match:
                signals.append(SignalRecord(current.key, current.raw_id, current.can_id, current.frame_format, match.group("name"), line_number))
    return messages, signals


def _mapping_base(mapping: Dict[str, Any], key: str, current_name: str) -> Tuple[str, Optional[str]]:
    entry = mapping.get(key)
    if not isinstance(entry, dict):
        return current_name, None
    original = str(entry.get("original_name", ""))
    generated = str(entry.get("generated_name", ""))
    if original and current_name in {original, generated}:
        return original, None
    if original and current_name != original:
        return current_name, "当前名称与已保存映射不一致，来源不明；请先确认原始名称。"
    return current_name, None


def _looks_generated(name: str, identifier: str, can_id: int, object_type: str) -> bool:
    can_hex = re.escape(format_can_id(can_id))
    ident = re.escape(identifier)
    if object_type == "message":
        return re.search(rf"_{ident}_{can_hex}$", name) is not None
    # 兼容本工具旧版无分隔符的名称，只生成当前要求的 ID 后下划线格式。
    return re.match(rf"^{ident}_sig{can_hex}(?:_)?", name) is not None


def build_rename_plan(path: str, config: Dict[str, Any]) -> RenamePlan:
    content, _encoding = _read_text(path)
    lines = content.splitlines(keepends=True)
    messages, signals = _records(lines)
    default_identifier = validate_identifier(config.get("default_identifier", ""))
    targets = config.get("targets", {})
    mapping = config.get("mapping", {})
    mapped_messages = mapping.get("messages", {}) if isinstance(mapping, dict) else {}
    mapped_signals = mapping.get("signals", {}) if isinstance(mapping, dict) else {}
    items: List[RenamePreviewItem] = []

    for record in messages:
        identifier = validate_identifier(config.get("message_overrides", {}).get(record.key, default_identifier))
        base, map_error = _mapping_base(mapped_messages, record.key, record.name)
        if map_error:
            status, reason = "阻塞", map_error
        elif not mapped_messages.get(record.key) and _looks_generated(record.name, identifier, record.can_id, "message"):
            status, reason = "阻塞", "名称带有类似生成格式但没有本工具映射，不能猜测原始报文名。"
        elif not targets.get("messages", True):
            status, reason = "跳过", "未启用报文重命名。"
        else:
            status, reason = "可执行", "真实报文 CAN ID；保持标准/扩展帧身份。"
        new_name = f"{base}_{identifier}_{format_can_id(record.can_id)}"
        if status == "可执行" and base == record.name and new_name == record.name:
            status, reason = "无需修改", "名称已经符合当前配置。"
        items.append(RenamePreviewItem("报文", record.key, record.raw_id, record.can_id, record.frame_format, record.name, new_name, identifier, status, reason))

    message_identifier = {record.key: validate_identifier(config.get("message_overrides", {}).get(record.key, default_identifier)) for record in messages}
    for record in signals:
        key = signal_key(record.message_key, record.name)
        mapping_key = key
        if mapping_key not in mapped_signals:
            for candidate_key, candidate in mapped_signals.items():
                if candidate_key.startswith(record.message_key + "|") and isinstance(candidate, dict) and candidate.get("generated_name") == record.name:
                    mapping_key = candidate_key
                    break
        override = config.get("signal_overrides", {}).get(mapping_key, config.get("signal_overrides", {}).get(key, message_identifier.get(record.message_key, default_identifier)))
        identifier = validate_identifier(override)
        base, map_error = _mapping_base(mapped_signals, mapping_key, record.name)
        if map_error:
            status, reason = "阻塞", map_error
        elif not mapped_signals.get(key) and _looks_generated(record.name, identifier, record.can_id, "signal"):
            status, reason = "阻塞", "名称带有类似生成格式但没有本工具映射，不能猜测原始信号名。"
        elif not targets.get("signals", True):
            status, reason = "跳过", "未启用信号重命名。"
        else:
            status, reason = "可执行", "信号继承所属报文标识；CAN ID 自动来自所属报文。"
        new_name = f"{identifier}_sig{format_can_id(record.can_id)}_{base}"
        if status == "可执行" and base == record.name and new_name == record.name:
            status, reason = "无需修改", "名称已经符合当前配置。"
        items.append(RenamePreviewItem("信号", mapping_key, record.raw_id, record.can_id, record.frame_format, record.name, new_name, identifier, status, reason))

    # 报文名全文件唯一；信号名在各自报文内唯一。
    for object_type, scope_items in (("报文", [i for i in items if i.object_type == "报文"]),):
        seen: Dict[str, RenamePreviewItem] = {}
        for item in scope_items:
            if item.status not in {"可执行", "无需修改"}:
                continue
            previous = seen.get(item.new_name)
            if previous and previous.object_key != item.object_key:
                item.status = "阻塞"
                item.reason = f"新名称与 {previous.old_name} 冲突。"
                if previous.status != "阻塞":
                    previous.status = "阻塞"
                    previous.reason = f"新名称与 {item.old_name} 冲突。"
            else:
                seen[item.new_name] = item
    by_message: Dict[str, Dict[str, RenamePreviewItem]] = {}
    for item in (i for i in items if i.object_type == "信号"):
        by_message.setdefault(item.object_key.split("|", 1)[0], {})
        if item.status not in {"可执行", "无需修改"}:
            continue
        scope = by_message[item.object_key.split("|", 1)[0]]
        previous = scope.get(item.new_name)
        if previous and previous.object_key != item.object_key:
            item.status = "阻塞"
            item.reason = f"同一报文内新名称与 {previous.old_name} 冲突。"
            if previous.status != "阻塞":
                previous.status = "阻塞"
                previous.reason = f"同一报文内新名称与 {item.old_name} 冲突。"
        else:
            scope[item.new_name] = item

    return RenamePlan(path, file_fingerprint(path), items, tuple(messages), tuple(signals))


def _replace_first_token(line: str, prefix_pattern: str, old_name: str, new_name: str) -> str:
    match = re.match(prefix_pattern, line)
    if not match or match.group(1) != old_name:
        return line
    return f"{line[:match.start(1)]}{new_name}{line[match.end(1):]}"


def _replace_members(line: str, replacements: Dict[str, str]) -> str:
    if ":" not in line:
        return line
    head, tail = line.split(":", 1)
    def repl(match: re.Match[str]) -> str:
        token = match.group(1)
        return replacements.get(token, token)
    return head + ":" + re.sub(r"(?<!\S)([^\s;]+)(?=[\s;]|$)", repl, tail)


def _replace_structured_references(
    lines: List[str],
    message_changes: Dict[str, Tuple[str, str]],
    signal_changes: Dict[str, Tuple[str, str]],
) -> List[str]:
    current_key: Optional[str] = None
    output: List[str] = []

    def signal_map_for(message_key: Optional[str]) -> Dict[str, str]:
        if not message_key:
            return {}
        return {
            old: new
            for (msg_key, old), (_old, new) in signal_changes.items()
            if msg_key == message_key
        }

    for line in lines:
        plain = _line_without_eol(line)
        eol = line[len(plain):]
        bo = BO_RE.match(plain)
        if bo:
            raw_id = int(bo.group("raw_id"))
            current_key = object_key(raw_id)
            old, new = message_changes.get(current_key, ("", ""))
            if old and bo.group("name").strip() == old:
                leading = bo.group("name")[: len(bo.group("name")) - len(bo.group("name").lstrip())]
                trailing = bo.group("name")[len(bo.group("name").rstrip()):]
                line = f"{bo.group('prefix')}{bo.group('raw_id')}{bo.group('middle')}{leading}{new}{trailing}{bo.group('suffix')}" + line[len(plain):]
            output.append(line)
            continue

        signal_map = signal_map_for(current_key)

        # SG_ 定义：只改第一个信号名 token，不动接收节点、单位和其它文本。
        if re.match(r"^\s*SG_\s+", plain) and not re.match(r"^\s*SG_(?:MUL_VAL|VALTYPE|GROUP)_", plain):
            signal_match = re.match(r"^(\s*SG_\s+)(\S+)(.*)$", plain)
            if signal_match and signal_match.group(2) in signal_map:
                line = f"{signal_match.group(1)}{signal_map[signal_match.group(2)]}{signal_match.group(3)}{eol}"
        elif re.match(r"^\s*CM_\s+SG_\s+\d+\s+\S+", plain):
            match = re.match(r"^(\s*CM_\s+SG_\s+)(\d+)(\s+)(\S+)(.*)$", plain)
            reference_map = signal_map_for(object_key(int(match.group(2)))) if match else {}
            if match and match.group(4) in reference_map:
                line = f"{match.group(1)}{match.group(2)}{match.group(3)}{reference_map[match.group(4)]}{match.group(5)}{eol}"
        elif re.match(r"^\s*BA_\s+\"[^\"]+\"\s+SG_\s+\d+\s+\S+", plain):
            match = re.match(r"^(\s*BA_\s+\"[^\"]+\"\s+SG_\s+)(\d+)(\s+)(\S+)(.*)$", plain)
            reference_map = signal_map_for(object_key(int(match.group(2)))) if match else {}
            if match and match.group(4) in reference_map:
                line = f"{match.group(1)}{match.group(2)}{match.group(3)}{reference_map[match.group(4)]}{match.group(5)}{eol}"
        elif re.match(r"^\s*(?:VAL_|SIG_VALTYPE_|SG_MUL_VAL_)\s+\d+\s+\S+", plain):
            match = re.match(r"^(\s*(?:VAL_|SIG_VALTYPE_|SG_MUL_VAL_)\s+)(\d+)(\s+)(\S+)(.*)$", plain)
            reference_map = signal_map_for(object_key(int(match.group(2)))) if match else {}
            if match and match.group(4) in reference_map:
                line = f"{match.group(1)}{match.group(2)}{match.group(3)}{reference_map[match.group(4)]}{match.group(5)}{eol}"
        elif re.match(r"^\s*SIG_GROUP_\s+\d+\s+\S+", plain):
            group_id = re.match(r"^\s*SIG_GROUP_\s+(\d+)", plain)
            line = _replace_members(plain, signal_map_for(object_key(int(group_id.group(1))) if group_id else None)) + eol
        output.append(line)
    return output


def _atomic_write(path: Path, content: str, encoding: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix=f".{path.name}.", suffix=".tmp", dir=str(path.parent))
    os.close(fd)
    temp_path = Path(temp_name)
    try:
        temp_path.write_text(content, encoding=encoding, newline="")
        os.replace(temp_path, path)
    finally:
        temp_path.unlink(missing_ok=True)


def build_dbc_repair_plan(path: str) -> DbcRepairPlan:
    """识别没有可用 BA_DEF_ 的属性赋值，绝不凭属性名编造定义。"""
    content, _encoding = _read_text(path)
    lines = content.splitlines(keepends=True)
    definitions: set[Tuple[str, str]] = set()
    for raw_line in lines:
        match = ATTRIBUTE_DEFINITION_RE.match(_line_without_eol(raw_line))
        if match:
            scope = (match.group("scope") or "GLOBAL").rstrip("_").upper()
            definitions.add((scope, match.group("name").lower()))

    items: List[DbcRepairItem] = []
    for line_number, raw_line in enumerate(lines, start=1):
        match = ATTRIBUTE_ASSIGNMENT_RE.match(_line_without_eol(raw_line))
        if not match:
            continue
        scope = (match.group("scope") or "GLOBAL").rstrip("_").upper()
        name = match.group("name")
        if (scope, name.lower()) in definitions or ("GLOBAL", name.lower()) in definitions:
            continue
        items.append(DbcRepairItem(
            rule_id="DBC_ATTR_UNDEFINED_001",
            line_number=line_number,
            attribute_name=name,
            scope=scope,
            raw_line=_line_without_eol(raw_line),
            status="待用户配置",
            reason=(
                f"属性“{name}”没有适用于 {scope} 的 BA_DEF_ 定义；"
                "无法可靠推断类型、枚举槽位或默认值。"
            ),
        ))
    return DbcRepairPlan(path, file_fingerprint(path), items)


def apply_dbc_repair_plan(
    plan: DbcRepairPlan,
    output_path: str,
    *,
    remove_undefined_attributes: bool = False,
) -> Tuple[str, Tuple[DbcRepairItem, ...]]:
    """将用户确认的“删除无定义属性”操作写入另存副本。

    删除未定义 BA_ 赋值会丢失该属性的业务含义，因此必须由调用方显式传入
    ``remove_undefined_attributes=True``。原 DBC 不会被覆盖。
    """
    if file_fingerprint(plan.source_path) != plan.source_fingerprint:
        raise ValueError("DBC 文件在修复预览后已发生变化；旧预览已失效，请重新检查。")
    if not plan.items:
        raise ValueError("没有检测到可处理的未定义属性赋值。")
    if not remove_undefined_attributes:
        raise ValueError("未定义属性缺少可靠定义；请提供 BA_DEF_，或明确确认删除这些 BA_ 赋值。")

    source = Path(plan.source_path)
    target = Path(output_path)
    if target.resolve() == source.resolve():
        raise ValueError("属性修复默认另存为新 DBC，避免直接覆盖原文件。")
    content, encoding = _read_text(plan.source_path)
    remove_lines = {item.line_number for item in plan.items}
    repaired = "".join(
        line for line_number, line in enumerate(content.splitlines(keepends=True), start=1)
        if line_number not in remove_lines
    )
    _atomic_write(target, repaired, encoding)

    remaining = build_dbc_repair_plan(str(target)).items
    if remaining:
        raise ValueError("修复后仍存在未定义属性赋值，已停止报告成功。")
    return str(target), tuple(plan.items)


def apply_dbc_attribute_definition(
    plan: DbcRepairPlan,
    item: DbcRepairItem,
    definition_tail: str,
    output_path: str,
) -> str:
    """根据用户提供的类型/范围补入 BA_DEF_，不猜测未知属性的语义。"""
    if item not in plan.items:
        raise ValueError("属性修复项不属于当前预览。")
    if file_fingerprint(plan.source_path) != plan.source_fingerprint:
        raise ValueError("DBC 文件在预览后已发生变化；请重新检查。")
    tail = definition_tail.strip().rstrip(";").strip()
    scope_prefix = "" if item.scope == "GLOBAL" else f"{item.scope}_ "
    definition = f'BA_DEF_ {scope_prefix}"{item.attribute_name}" {tail};'
    if not ATTRIBUTE_DEFINITION_RE.match(definition):
        raise ValueError('属性定义格式不正确。例如：INT 0 255，或 ENUM "No","Yes"。')
    source = Path(plan.source_path)
    target = Path(output_path)
    if target.resolve() == source.resolve():
        raise ValueError("补充属性定义默认另存为新 DBC，避免直接覆盖原文件。")
    content, encoding = _read_text(plan.source_path)
    lines = content.splitlines(keepends=True)
    newline = "\r\n" if "\r\n" in content else "\n"
    lines.insert(item.line_number - 1, definition + newline)
    _atomic_write(target, "".join(lines), encoding)
    remaining = [
        candidate for candidate in build_dbc_repair_plan(str(target)).items
        if candidate.attribute_name.lower() == item.attribute_name.lower() and candidate.scope == item.scope
    ]
    if remaining:
        raise ValueError("新增 BA_DEF_ 后属性仍未通过复检，已停止报告成功。")
    return str(target)


def apply_rename_plan(plan: RenamePlan, config: Dict[str, Any], output_path: str) -> Tuple[str, Dict[str, Any]]:
    if file_fingerprint(plan.source_path) != plan.source_fingerprint:
        raise ValueError("DBC 文件在预览后已发生变化；旧预览已失效，请重新读取并预览。")
    if plan.blocked_items:
        details = "\n".join(f"{item.object_type} {item.old_name}：{item.reason}" for item in plan.blocked_items[:12])
        raise ValueError(f"存在不可执行的命名项，已阻止写回：\n{details}")
    content, encoding = _read_text(plan.source_path)
    lines = content.splitlines(keepends=True)
    message_changes = {item.object_key: (item.old_name, item.new_name) for item in plan.items if item.object_type == "报文" and item.status == "可执行"}
    signal_changes = {
        (item.object_key.split("|", 1)[0], item.old_name): (item.old_name, item.new_name)
        for item in plan.items if item.object_type == "信号" and item.status == "可执行"
    }
    changed_lines = _replace_structured_references(lines, message_changes, signal_changes)
    target = Path(output_path)
    source = Path(plan.source_path)
    if target.resolve() == source.resolve():
        raise ValueError("默认另存为修正后的 DBC；如需覆盖原文件，请先明确传入覆盖流程。")
    _atomic_write(target, "".join(changed_lines), encoding)

    mapping = config.setdefault("mapping", {"source_fingerprint": "", "messages": {}, "signals": {}})
    mapping["source_fingerprint"] = file_fingerprint(plan.source_path)
    mapping["target_fingerprint"] = file_fingerprint(str(target))
    mapping.setdefault("messages", {})
    mapping.setdefault("signals", {})
    for item in plan.items:
        if item.status == "可执行":
            entry = {"original_name": item.old_name, "generated_name": item.new_name, "can_id": format_can_id(item.can_id), "frame_format": item.frame_format}
            if item.object_type == "报文":
                mapping["messages"][item.object_key] = entry
            else:
                mapping["signals"][item.object_key] = entry
    return str(target), config


@dataclass
class NodeCompletionPlan:
    source_path: str
    source_fingerprint: str
    referenced_nodes: Tuple[str, ...]
    declared_nodes: Tuple[str, ...]
    missing_nodes: Tuple[str, ...]


def build_node_completion_plan(path: str) -> NodeCompletionPlan:
    content, _encoding = _read_text(path)
    lines = content.splitlines(keepends=True)
    referenced: List[str] = []
    declared: List[str] = []
    for raw_line in lines:
        line = _line_without_eol(raw_line)
        bu = re.match(r"^\s*BU_\s*:\s*(.*)$", line)
        if bu:
            declared.extend(token for token in bu.group(1).split() if token != "Vector__XXX")
            continue
        bo = BO_RE.match(line)
        if bo:
            sender_match = re.search(r":\s*\d+\s+(\S+)", bo.group("suffix"))
            if sender_match and sender_match.group(1) != "Vector__XXX":
                referenced.append(sender_match.group(1))
            continue
        if re.match(r"^\s*SG_\s+", line) and not re.match(r"^\s*SG_(?:MUL_VAL|VALTYPE|GROUP)_", line):
            tail = line.split('"', 1)[-1] if '"' in line else ""
            # 语句最后的接收节点列表位于单位字符串之后；只取最后一个字段，避免把单位当节点。
            sg = re.match(r"^\s*SG_\s+\S+.*?\"[^\"]*\"\s*(.*)$", line)
            if sg:
                referenced.extend(token for token in sg.group(1).split() if token != "Vector__XXX")
    unique_declared = tuple(dict.fromkeys(declared))
    unique_referenced = tuple(dict.fromkeys(referenced))
    missing = tuple(node for node in unique_referenced if node not in set(unique_declared))
    return NodeCompletionPlan(path, file_fingerprint(path), unique_referenced, unique_declared, missing)


def apply_node_completion(plan: NodeCompletionPlan, output_path: str) -> Tuple[str, Tuple[str, ...]]:
    if file_fingerprint(plan.source_path) != plan.source_fingerprint:
        raise ValueError("DBC 文件在预览后已发生变化；节点补全预览已失效，请重新读取。")
    if not plan.missing_nodes:
        raise ValueError("没有可补入 BU_ 的真实节点。")
    content, encoding = _read_text(plan.source_path)
    lines = content.splitlines(keepends=True)
    newline = "\r\n" if "\r\n" in content else "\n"
    target = Path(output_path)
    if target.resolve() == Path(plan.source_path).resolve():
        raise ValueError("节点补全默认另存为新 DBC，避免直接覆盖原文件。")
    inserted = False
    for index, raw_line in enumerate(lines):
        if re.match(r"^\s*BU_\s*:", _line_without_eol(raw_line)):
            end = "\r\n" if raw_line.endswith("\r\n") else "\n" if raw_line.endswith("\n") else ""
            base = _line_without_eol(raw_line).rstrip()
            lines[index] = f"{base} {' '.join(plan.missing_nodes)}{end}"
            inserted = True
            break
    if not inserted:
        position = next((i for i, raw_line in enumerate(lines) if re.match(r"^\s*BO_\s+", _line_without_eol(raw_line))), len(lines))
        lines.insert(position, f"BU_: {' '.join(plan.declared_nodes + plan.missing_nodes)}{newline}")
    _atomic_write(target, "".join(lines), encoding)
    return str(target), plan.missing_nodes
