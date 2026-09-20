#!/usr/bin/env python3
"""DBC命名和节点补全的 focused 回归测试。"""

from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from dbc_transform import (
    apply_node_completion,
    apply_rename_plan,
    build_node_completion_plan,
    build_rename_plan,
    default_rename_config,
    load_rename_config,
    save_rename_config,
)
import can_matrix_checker as checker


DBC_TEXT = """VERSION \"\"\n
NS_ :\n
    CM_\n
    BA_DEF_\n
    BA_\n
    VAL_\n
    SIG_GROUP_\n
    SIG_VALTYPE_\n
    SG_MUL_VAL_\n
BS_:\n
BU_: TX\n
BO_ 257 VehicleStatus: 8 TX\n
 SG_ DataLenght : 0|8@1+ (1,0) [0|255] \"\" RX\n
 SG_ KeepSignal : 8|8@1+ (1,0) [0|255] \"\" RX\n
BO_ 562 OtherStatus: 8 TX\n
 SG_ DataLenght : 0|8@1+ (1,0) [0|255] \"\" RX\n
CM_ SG_ 257 DataLenght \"DataLenght\";\n
BA_ \"GenSigStartValue\" SG_ 257 DataLenght 0;\n
VAL_ 257 DataLenght 0 \"Closed\";\n
SIG_GROUP_ 257 DataGroup 1 : DataLenght;\n
SIG_VALTYPE_ 257 DataLenght 1;\n
SG_MUL_VAL_ 257 DataLenght 1 0;\n"""


class DbcTransformTests(unittest.TestCase):
    def test_real_id_naming_is_idempotent_and_updates_references(self) -> None:
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.dbc"
            first = Path(tmp) / "first.dbc"
            second = Path(tmp) / "second.dbc"
            source.write_text(DBC_TEXT, encoding="utf-8")
            config = default_rename_config()
            plan = build_rename_plan(str(source), config)
            message_items = [item for item in plan.items if item.object_type == "报文"]
            signal_items = [item for item in plan.items if item.object_type == "信号"]
            self.assertEqual(message_items[0].new_name, "VehicleStatus_can1_0x101")
            signal_names = {(item.can_id, item.old_name): item.new_name for item in signal_items}
            self.assertEqual(signal_names[(0x101, "DataLenght")], "can1_sig0x101DataLenght")
            self.assertEqual(signal_names[(0x101, "KeepSignal")], "can1_sig0x101KeepSignal")
            self.assertEqual(message_items[1].new_name, "OtherStatus_can1_0x232")
            self.assertEqual(signal_names[(0x232, "DataLenght")], "can1_sig0x232DataLenght")
            apply_rename_plan(plan, config, str(first))
            first_text = first.read_text(encoding="utf-8")
            self.assertIn("VehicleStatus_can1_0x101", first_text)
            self.assertIn("can1_sig0x232DataLenght", first_text)
            self.assertIn("SG_ can1_sig0x101KeepSignal :", first_text)
            self.assertIn("CM_ SG_ 257 can1_sig0x101DataLenght", first_text)
            self.assertIn("VAL_ 257 can1_sig0x101DataLenght", first_text)
            self.assertIn("SIG_GROUP_ 257 DataGroup 1 : can1_sig0x101DataLenght", first_text)
            original_db = checker.parse_dbc(str(source))
            renamed_db = checker.parse_dbc(str(first))
            compared = checker.compare_databases(original_db, renamed_db, [], vector_rules_enabled=False, rename_mapping=config)
            self.assertFalse(any(item.category == "缺失信号" for item in compared))

            config["default_identifier"] = "can2"
            plan2 = build_rename_plan(str(first), config)
            self.assertEqual(next(i.new_name for i in plan2.items if i.object_type == "信号" and i.old_name == "can1_sig0x101DataLenght"), "can2_sig0x101DataLenght")
            apply_rename_plan(plan2, config, str(second))
            second_text = second.read_text(encoding="utf-8")
            self.assertNotIn("can1_sig0x101can1_sig0x101", second_text)
            self.assertIn("can2_sig0x101DataLenght", second_text)

    def test_config_roundtrip_and_node_completion(self) -> None:
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "source.dbc"
            target = Path(tmp) / "nodes.dbc"
            config_path = Path(tmp) / "rename.json"
            source.write_text(DBC_TEXT.replace("BU_: TX", "BU_: TX"), encoding="utf-8")
            config = default_rename_config()
            config["default_identifier"] = "CAN1"
            save_rename_config(str(config_path), config)
            self.assertEqual(load_rename_config(str(config_path))["default_identifier"], "CAN1")
            node_plan = build_node_completion_plan(str(source))
            self.assertEqual(node_plan.missing_nodes, ("RX",))
            apply_node_completion(node_plan, str(target))
            self.assertIn("BU_: TX RX", target.read_text(encoding="utf-8"))

    def test_standard_and_extended_same_numeric_id_keep_distinct_identity(self) -> None:
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "frames.dbc"
            source.write_text(
                DBC_TEXT.replace("BO_ 562 OtherStatus", "BO_ 2147483905 OtherStatus"),
                encoding="utf-8",
            )
            plan = build_rename_plan(str(source), default_rename_config())
            messages = [item for item in plan.items if item.object_type == "报文"]
            self.assertEqual({item.object_key for item in messages}, {"standard:0x101", "extended:0x101"})
            self.assertEqual({item.new_name for item in messages}, {"VehicleStatus_can1_0x101", "OtherStatus_can1_0x101"})

    def test_empty_bu_still_reports_real_node_references(self) -> None:
        with TemporaryDirectory() as tmp:
            source = Path(tmp) / "missing_nodes.dbc"
            source.write_text(
                'VERSION ""\nBS_:\nBU_:\nBO_ 257 Msg: 8 TX\n SG_ Sig : 0|8@1+ (1,0) [0|255] "" RX\n',
                encoding="utf-8",
            )
            db = checker.parse_dbc(str(source))
            rule_ids = {item.rule_id for item in checker.self_check_database(db, "DBC")}
            self.assertIn("DBC_NODE_REF_001", rule_ids)
            self.assertIn("DBC_NODE_REF_002", rule_ids)


if __name__ == "__main__":
    unittest.main()
