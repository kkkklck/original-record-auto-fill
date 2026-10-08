<div align="center">

<img src="docs/assets/cover.svg" alt="原始记录自动填写 — Word 测点数据到 Excel 报告与 Word 汇总" width="100%">

# 原始记录自动填写 · The Unification

**防火检测数据整理 · 构件分类 · 自动分页 · Excel / Word 输出**

Turn Word inspection tables into organized Excel records and a Word summary.

![Python](https://img.shields.io/badge/Python-3.10%2B-3677a9?style=flat-square&logo=python&logoColor=white) ![Word](https://img.shields.io/badge/Input-DOCX-365f88?style=flat-square) ![Excel](https://img.shields.io/badge/Output-XLSX%20%2B%20DOCX-36735e?style=flat-square) [![License](https://img.shields.io/badge/License-Commercial-c68b49?style=flat-square)](LICENSE)

[功能概览](#功能概览) · [四种工作模式](#四种工作模式) · [快速开始](#快速开始) · [使用指南](docs/USAGE.md) · [问题反馈](https://github.com/kkkklck/original-record-auto-fill/issues)

</div>

---

面向**钢结构防火检测原始记录**的数据整理工具。读取 Word 中的测点表格，识别钢柱、钢梁、支撑、网架等构件，按日期或楼层分配数据，再填入 Excel 模板并生成 Word 汇总，方便核对与归档。

## 功能概览

| 能力 | 工作内容 |
| :--- | :--- |
| **Word 表格读取** | 识别包含“测点1”“平均值”的数据表，处理相关合并单元格 |
| **构件分类与排序** | 区分钢柱、钢梁、支撑、网架与其他构件，按楼层或编号整理 |
| **按规则分配** | 支持日期分桶、楼层断点、单日出表与楼层 × 日期切片 |
| **模板填写与分页** | 复制所需工作表、填写元信息，分开处理普通页与 μ 页 |
| **报告与汇总** | 输出 Excel 报告与 Word 汇总；Excel 同名输出自动追加序号 |
| **交互式引导** | 命令行逐步提示，支持查看帮助、返回上一步与分配预览 |

## 从输入到交付

```mermaid
flowchart LR
    A[Word 测点表格] --> B[读取与构件分类]
    B --> C[选择模式与分配规则]
    C --> D[Excel 模板自动填写]
    B --> E[Word 汇总]
    D --> F[核对与归档]
    E --> F
```

## 四种工作模式

| 模式 | 适合的场景 | 分配方式 |
| :--- | :--- | :--- |
| **Mode 1 · 按日期** | 同批构件分多天出表 | 配置日期与构件规则，预览后确认 |
| **Mode 2 · 按楼层断点** | 不同楼层区间对应不同日期 | 用断点划分楼层，为每桶设置日期 |
| **Mode 3 · 单日** | 整单使用同一日期 | 填一次日期与温度，自动整理分页 |
| **Mode 4 · 楼层 × 日期** | 同一楼层跨多天安排 | 按计划均分或设置每日上限 |

第一次体验可从 **Mode 3** 开始。各模式的输入规则见 [使用指南](docs/USAGE.md)。

## 快速开始

### 1. 获取项目与依赖

源码使用 `str | None`、`list[int]` 等类型注解，建议使用 **Python 3.10+**。

```bash
git clone https://github.com/kkkklck/original-record-auto-fill.git
cd original-record-auto-fill
python -m pip install openpyxl python-docx
```

### 2. 指定 Excel 模板

打开 `Original record auto-fill program.py`，找到文件开头的 `XLSX_WITH_SUPPORT_DEFAULT`，将其设置为你本机的模板路径。仓库附带的模板为 **`防火excel模板μ.xlsx`**。

例如，在 Windows 上将配置改为实际绝对路径：

```python
XLSX_WITH_SUPPORT_DEFAULT = Path(r"C:\your-folder\original-record-auto-fill\防火excel模板μ.xlsx")
```

当前源码默认值是作者本机路径，直接在另一台电脑运行前需要调整。模板工作表名称与结构应和程序匹配；涉及 μ 数据时也要保留相应 μ 模板页。

### 3. 运行并生成记录

```bash
python "Original record auto-fill program.py"
```

1. 输入 Word `.docx` 路径，可先使用仓库内的 `示例.docx`。
2. 按提示填写工程名称、委托编号等信息。
3. 选择工作模式，填写日期、温度及相关分配规则。
4. 核对分配预览，确认后生成 Excel 报告。
5. 对照 Word 汇总检查构件、数据、日期与分页结果。

### 输出文件

| 文件 | 保存位置与用途 |
| :--- | :--- |
| `The Unification_报告版.xlsx` | Word 源文件同目录，用于查看与打印原始记录；重名时追加序号 |
| `汇总原始记录.docx` | Word 源文件同目录，用于核对提取数据；再次运行前按需备份已有汇总 |

## 使用提示

<details>
<summary><strong>输入要求与快捷指令</strong></summary>

Word 数据需保存在 `.docx` 表格中，相关表头包含“测点1”“平均值”。运行前关闭涉及的 Word、Excel 文件，避免文件被占用。

| 输入 | 作用 |
| :--- | :--- |
| `help` | 在路径输入界面打开帮助 |
| `q` | 在交互步骤返回上一步 |
| `Q` | 在路径输入界面退出程序 |

更详细的日期、编号范围与模式说明见 [使用指南](docs/USAGE.md)。

</details>

<details>
<summary><strong>常见问题</strong></summary>

| 现象 | 处理建议 |
| :--- | :--- |
| 提示模板不存在 | 修改 `XLSX_WITH_SUPPORT_DEFAULT` 为本机实际路径 |
| 未识别到数据 | 检查 `.docx` 表格与“测点1”“平均值”表头 |
| Excel 无法保存 | 关闭正在打开的模板、目标报告后重试 |
| 提示缺少 μ 模板 | 检查对应类别的 μ 工作表是否存在 |
| 构件分类或日期不符合预期 | 先核对构件名称与分配预览，再调整规则 |

</details>

## 项目结构

```text
original-record-auto-fill/
├── Original record auto-fill program.py   # 命令行程序入口
├── 防火excel模板μ.xlsx                     # Excel 模板
├── 示例.docx                              # 示例 Word 输入
├── LICENSE                                # 商业许可条款
└── docs/                                  # 主页素材与使用指南
```

## 许可与反馈

本项目采用 **商业许可**，使用需取得作者授权，具体范围以 [LICENSE](LICENSE) 与授权约定为准。可通过 [作者 GitHub 主页](https://github.com/kkkklck) 联系作者，通过 [Issues](https://github.com/kkkklck/original-record-auto-fill/issues) 反馈问题。

<div align="center">

<sub>Built by LCK · Less copying. More clarity.</sub>

</div>
