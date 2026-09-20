#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CAN矩阵-DBC检查工具 v3.0 离线自测。"""

from pathlib import Path

import can_matrix_checker as checker

BASE = Path(__file__).resolve().parent


def main() -> None:
    rules = checker.load_rules(str(BASE / "can_rules.json"))

    matrix, _, _ = checker.read_matrix(str(BASE / "sample_matrix.csv"))
    dbc = checker.parse_dbc(str(BASE / "sample.dbc"))
    normal_diffs = checker.compare_databases(matrix, dbc, rules, vector_rules_enabled=False)
    assert not normal_diffs, f"基础样例应为0项，实际：{normal_diffs}"

    semantic_matrix, _, _ = checker.read_matrix(str(BASE / "sample_semantic_matrix.csv"))
    semantic_dbc = checker.parse_dbc(str(BASE / "sample_semantic_nm.dbc"))
    semantic_diffs = checker.compare_databases(semantic_matrix, semantic_dbc, rules, vector_rules_enabled=False)
    assert any(d.rule_id == "NM_PROJECT_001" for d in semantic_diffs), "NM项目规则未命中"

    project_matrix, _, _ = checker.read_matrix(str(BASE / "sample_project_rules_matrix.csv"))
    project_dbc = checker.parse_dbc(str(BASE / "sample_project_rules_bad.dbc"))
    project_diffs = checker.compare_databases(project_matrix, project_dbc, rules, vector_rules_enabled=False)
    expected_rule_ids = {
        "NM_PROJECT_001",
        "XCP_PROJECT_001",
        "DBC_CANFD_DLC_001",
        "DBC_TX_EXPLICIT_001",
        "DBC_TX_EXPLICIT_002",
        "DBC_NODE_REF_002",
        "DBC_ATTR_UNDEFINED_001",
    }
    actual_rule_ids = {d.rule_id for d in project_diffs}
    missing = expected_rule_ids - actual_rule_ids
    assert not missing, f"项目准出规则缺失：{sorted(missing)}；实际={sorted(actual_rule_ids)}"

    # v2.5：不提供矩阵也应能执行同一批DBC自身/项目规则。
    dbc_only_diffs = checker.check_dbc_only(project_dbc, rules, vector_rules_enabled=False)
    dbc_only_rule_ids = {d.rule_id for d in dbc_only_diffs}
    missing_dbc_only = expected_rule_ids - dbc_only_rule_ids
    assert not missing_dbc_only, f"无矩阵模式规则缺失：{sorted(missing_dbc_only)}"

    report = BASE / "self_test_report.xlsx"
    checker.export_xlsx_report(str(report), project_diffs, project_matrix, project_dbc)
    assert report.is_file() and report.stat().st_size > 0, "矩阵报告导出失败"
    report.unlink()

    dbc_only_report = BASE / "self_test_dbc_only_report.xlsx"
    checker.export_xlsx_report(str(dbc_only_report), dbc_only_diffs, None, project_dbc)
    assert dbc_only_report.is_file() and dbc_only_report.stat().st_size > 0, "无矩阵报告导出失败"
    dbc_only_report.unlink()

    sheets, best_sheet, _details = checker.inspect_matrix_sheets(str(BASE / "sample_report_v2.xlsx"))
    assert sheets and best_sheet is None, "检查报告不应被识别为原始通信矩阵"
    try:
        checker.validate_matrix_selection(str(BASE / "sample_report_v2.xlsx"), "检查汇总")
    except checker.MatrixReadError:
        pass
    else:
        raise AssertionError("检查报告应在开始检查前被拒绝")

    # 不同简单复用分支可以使用同一bit，不能误报重叠。
    mux_dbc_path = BASE / "_self_test_mux.dbc"
    mux_dbc_path.write_text(
        'VERSION ""\nNS_ :\nBS_:\nBU_: ECU RX\n'
        'BO_ 256 MuxMsg: 8 ECU\n'
        ' SG_ Mux M : 0|4@1+ (1,0) [0|15] "" RX\n'
        ' SG_ DataA m1 : 8|8@1+ (1,0) [0|255] "" RX\n'
        ' SG_ DataB m2 : 8|8@1+ (1,0) [0|255] "" RX\n',
        encoding="utf-8",
    )
    try:
        mux_db = checker.parse_dbc(str(mux_dbc_path))
        mux_diffs = checker.check_dbc_only(mux_db, [], vector_rules_enabled=False)
        assert not any(d.rule_id == "DBC_SIG_OVERLAP_001" for d in mux_diffs), "互斥复用分支不应报重叠"
    finally:
        mux_dbc_path.unlink(missing_ok=True)

    # v3.0：Vector手册内置规则应命中标准属性与跨属性问题。
    vector_bad = checker.parse_dbc(str(BASE / "sample_vector_manual_bad.dbc"))
    vector_diffs = checker.check_dbc_only(vector_bad, [], vector_rules_enabled=True)
    vector_rule_ids = {d.rule_id for d in vector_diffs}
    expected_vector = {
        "VEC_GENERAL_BUSTYPE_001",
        "VEC_COM_ONCHANGE_SIZE_001",
        "VEC_COM_REPETITION_001",
        "VEC_UPDATE_LENGTH_001",
        "VEC_UPDATE_SENDTYPE_001",
        "VEC_XCP_LAYER_001",
        "VEC_XCP_PAYLOAD_001",
    }
    missing_vector = expected_vector - vector_rule_ids
    assert not missing_vector, f"Vector手册规则缺失：{sorted(missing_vector)}；实际={sorted(vector_rule_ids)}"

    print("自测通过：矩阵对比、无矩阵检查、属性、NM/UDS/XCP、CAN FD、节点、复用和报告导出均正常。")


if __name__ == "__main__":
    main()
