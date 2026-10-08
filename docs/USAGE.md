# User guide · The Unification

[← Back to the project](../README.md)

## Input files

Use `.docx` source documents with measurement data in tables that contain the measurement-point and average-value headers recognized by the parser. Configure `XLSX_WITH_SUPPORT_DEFAULT` to point to the included `防火excel模板μ.xlsx` template on your machine.

Worksheet names and layout are part of the filling logic. Use a compatible template; renaming an unrelated workbook does not make it compatible.

## Choosing a mode

### Mode 1 · Date buckets

Distribute one batch across multiple dates. Enter dates and component rules. For braces or space frames, choose the identifier-based or floor-based strategy when prompted. Select an overlap priority if needed and review the allocation before generating records.

### Mode 2 · Floor breakpoints

Divide floors into intervals using breakpoints and assign dates and optional temperatures to the resulting buckets. Floor sorting includes basement floors, above-ground floors, equipment floors, and the roof.

### Mode 3 · Single date

Use one date for the entire batch, with optional temperature. The program paginates by component category and page capacity. This is a convenient starting point for a first run.

### Mode 4 · Floor × date

Distribute a floor's records across several dates with a shared plan or per-floor configuration. Use equal allocation or daily limits. A default plan can cover unconfigured floors; follow the prompts to handle remaining records.

## Dates and shortcuts

Supported date formats include `2026-10-08`, `2026/10/08`, and `20261008`. Dates without a year follow the program's normalization rules; check the resulting date before export.

| Input | Context and action |
| :--- | :--- |
| `help` | Opens help at the source-path prompt, including mode-specific instructions |
| `q` | Returns to the previous interaction step |
| `Q` | Exits at the source-path prompt |
| `*` | Accepts all items in range inputs that support this command |
| `lk` | Excludes a space-frame range where supported |
| `a` | Assigns unallocated components to the last date at the allocation-confirmation step |

Special commands depend on the current step. Follow the displayed instructions.

## Ordinary and μ pages

The program uses rules such as four-or-more-digit integers or absolute values of at least 1000 to classify μ data. Ordinary and μ pages are handled separately; ordinary pages precede μ pages within a bucket. Retain the corresponding μ template worksheets.

This is an input-classification rule. Check the actual units in your inspection records. The program fills instrument and other metadata according to its rules and configuration; confirm they match the equipment actually used.

## Checking the output

1. Check component names, categories, readings, and averages in the Word summary.
2. Check dates, temperatures, instruments, pagination, and ordinary/μ pages in Excel.
3. Confirm that unallocated items in the preview have been handled.
4. Confirm the files were saved successfully before proceeding.

Conflicting Excel names receive a sequence suffix. The Word summary uses a fixed filename, so back it up before rerunning if needed. Both outputs are saved beside the source Word document.

## Troubleshooting

- **Missing template:** set `XLSX_WITH_SUPPORT_DEFAULT` to an existing local path.
- **Locked file:** close the relevant Word and Excel files, then retry.
- **No data recognized:** check the input format, headers, and table layout.
- **Unexpected classification:** check component names against the recognition rules.
- **Missing μ sheets:** retain the appropriate category's μ template worksheets.

## License

Usage is governed by the [commercial license](../LICENSE) and your authorization agreement.
