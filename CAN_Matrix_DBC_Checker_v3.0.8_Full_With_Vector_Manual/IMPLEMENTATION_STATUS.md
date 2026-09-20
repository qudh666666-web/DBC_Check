# 本轮实现状态

## 已完成

- 建立根目录 Git 仓库，保留原始 ZIP、手册和样例；工作分支为 `feature/dbc-completion-naming`。
- 增加报文/信号命名预览与配置窗口：默认标识、报文覆盖、信号覆盖、多选批量设置、仅报文/仅信号处理、JSON保存/加载。
- 命名严格使用每条 `BO_` 的真实 CAN ID：`VehicleStatus_can1_0x101`、`can1_sig0x101DataLenght`；标准帧/扩展帧身份分开处理。
- 受控同步 `SG_`、`CM_ SG_`、`BA_ ... SG_`、`VAL_`、`SIG_GROUP_`、`SIG_VALTYPE_`、`SG_MUL_VAL_`，保留原文、注释、值表显示文字和未知语句。
- 通过原名→新名映射支持重复执行、`can1` 到 `can2` 替换和矩阵匹配，防止整批假缺失。
- 增加只补入明确引用真实节点的 `BU_` 节点补全入口；空 `BU_` 时现有节点检查仍报告真实收发节点引用问题。
- 写回默认另存新 DBC，预览校验输入指纹，临时文件原子写入并重新解析验证。

## 未完成 / 下一步

- 属性补全/修正的结构化修正项、属性来源展示、ENUM槽位核对和应用依赖尚未实现。
- 没有真实接收节点时的全局/报文/信号接收者配置入口尚未实现；当前不会虚构接收者。
- 对所有未支持 DBC 语法的名称引用阻塞扫描、Update Bit 客户E2E联动和更完整的保留语法覆盖仍需补充。
- 原 LibreOffice 嵌入式 Python 3.12.13 仍没有 `tkinter`；已安装用户范围官方 Python 3.12.10（Tk 8.6），并将 `run_checker.bat` 固定到该解释器。实际创建 `Tk()` 根窗口和短时启动主程序均通过；未进行人工点击级 GUI 验收。

## 测试结果

- `D:\Program Files\LibreOffice\program\python.exe -m py_compile can_matrix_checker.py dbc_transform.py test_dbc_transform.py`：通过。
- `D:\Program Files\LibreOffice\program\python.exe test_dbc_transform.py`：4 项通过，覆盖真实ID、重复命名、标准/扩展帧区分、引用同步、空 `BU_` 节点检查及矩阵映射。
- `D:\Program Files\LibreOffice\program\python.exe self_test.py`：通过；先安装了既有 `requirements.txt` 中的 `openpyxl` 依赖。
- `import tkinter`：失败，指定解释器缺少 tkinter；未伪装成 GUI 通过。
- `C:\Users\l\AppData\Local\Programs\Python\Python312\python.exe -c "import tkinter; ...; Tk()"`：通过，Tk 8.6。
- 使用该 Python 短时启动 `can_matrix_checker.py`：进程正常保持运行，随后结束冒烟进程。

## 相关文件与提交

- `can_matrix_checker.py`：主程序、GUI入口、节点检查和矩阵映射接入。
- `dbc_transform.py`：命名、映射、受控写回和节点补全核心逻辑。
- `test_dbc_transform.py`：本轮 focused 回归。
- `rename_config.example.json`：配置示例。
- 核心实现提交：`d975cf8`（Add DBC naming and node completion tools）。
- 状态与验证记录提交：`bdfba11`（Document implementation status and validation limits）。
- 原始基线：`b0585c9`（Initialize DBC checker baseline）。

## 启动方式

在源码目录运行：

```bat
run_checker.bat
```

也可以双击 `run_checker.bat`。命名配置可通过“命名设置”窗口保存到任意 JSON；应用后默认在生成 DBC 旁保存 `.rename.json` 映射配置。
