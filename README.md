# original record auto-fill program· Commercial Edition

A specialized Python automation suite for consolidating fireproofing measurement data.  
It extracts readings from inspection Word tables, classifies structural members, and
produces a polished Excel deliverable together with a verification-ready Word summary.
The latest iteration refines console guidance, improves template management, and
adds stricter validation for production use.

---

## Features
- **Guided console workflow** – step-by-step prompts with contextual hints,
  quick commands (`help`, `k`, `Q`), and graceful exit handling streamline daily use.
- **Robust document parsing** – reads `.docx` tables containing the keywords `"测点1"`
  and `"平均值"`, automatically normalizes merged columns, and highlights header rows.
- **Component intelligence** – recognizes steel columns, beams, braces, space frames,
  and falls back to **Others** with consistent formatting when data is ambiguous.
- **Excel automation** – fills the template `原始记录excel模板.xlsx`, prunes unused
  worksheets, keeps Times New Roman for the `μ` symbol, and chooses instrument models
  (`23-90` / `24-57`) from average values.
- **Progress & validation** – real-time progress indicators, minimum-row checks, and
  duplicate-name safeguards reduce manual QA effort.
- **Summary Word export** – generates `汇总原始记录.docx` alongside the source
  documents for quick spot checks and archiving.
- **Cross-platform paths** – supports Windows, macOS, and Linux; remember to close
  Word / Excel files before running the script to avoid locked file warnings.

---

## Repository Structure
```
.
├── Original record auto-fill program.py   # Main Python script
├── README.md                             # Documentation
├── LICENSE                               # Commercial license terms
├── 原始记录excel模板.xlsx                 # Excel template
└── 示例.docx                               # Sample Word input
```



---

## Installation
1. **Clone the repository**
   ```bash
   git clone https://github.com/your-username/the-unification-of-my-first-code-thinking.git
   cd the-unification-of-my-first-code-thinking
   ```
2. **Install Python 3.8+ (3.6 minimum)**
3. **Install dependencies**
   ```bash
   pip install openpyxl python-docx
   # If the network is slow in China:
   pip install openpyxl python-docx -i https://pypi.tuna.tsinghua.edu.cn/simple
   ```

---

## Usage
1. Place your Word measurement files and the Excel template in the working directory.
2. Run the script:
   ```bash
   python "Original record auto-fill program.py"
   ```
3. Follow the interactive prompts to provide source paths, component types, date
   buckets, and page counts.
4. The script outputs `The Unification_报告版.xlsx` (auto-incremented on conflicts)
   and `汇总原始记录.docx` next to the source Word file.

---

## Output Options
- `The Unification_报告版.xlsx`
- `汇总原始记录.docx`

---

## License
This project is distributed under a **Commercial License**. Usage requires a valid
paid license obtained from the author. See [LICENSE](LICENSE) for the full terms.
