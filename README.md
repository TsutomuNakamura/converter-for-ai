# converter-for-ai

Convert a folder of office documents, spreadsheets, slides, PDFs and text files into a small set of Markdown files that any AI assistant can read easily.

AI tools handle Markdown far better than binary formats like `.xlsx` or `.pptx`. `convert_to_markdown.py` walks a directory, extracts the text and tables from every supported file, and merges the results into numbered Markdown files, one series per file type.

## Supported formats

| Extension | How it is converted |
| --- | --- |
| `.docx` | Headings and paragraphs as text, tables as Markdown tables |
| `.doc` | Converted to `.docx` with LibreOffice, then handled as `.docx` |
| `.xlsx`, `.xlsm` | One section per sheet; tables as Markdown tables |
| `.csv` | Same as a single-sheet spreadsheet |
| `.pptx`, `.potx` | One section per slide with its text |
| `.pdf` | One section per page (text layer only) |
| `.txt` | Copied as-is |

Anything else is skipped, including images (`.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.svg`).

## Requirements

- Python 3 (tested with 3.14)
- Python packages: `python-docx`, `pandas`, `openpyxl`, `python-pptx`, `pypdf`, `tabulate`
- [LibreOffice](https://www.libreoffice.org/), only if you need to convert legacy `.doc` files. The `libreoffice` command must be on your `PATH`.

## Installation

```bash
git clone https://github.com/TsutomuNakamura/converter-for-ai.git
cd converter-for-ai

python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

> `tabulate` is not imported directly, but pandas needs it to write Markdown tables. Without it, spreadsheets and CSVs fail with `Import tabulate failed`. If you install packages by hand, use `python-docx`, not the unrelated `docx` package.

## Usage

Put the files you want to convert in a `resources/` folder (subfolders are searched recursively) and run the script from the directory that contains `resources/`. What it converts depends on the arguments:

```bash
# No arguments: convert every supported file under resources/
python convert_to_markdown.py

# One file: convert just that file
python convert_to_markdown.py budget.docx

# Several files: convert them and merge them into one Markdown file
python convert_to_markdown.py budget.docx sales.xlsx
```

Read the results in `merged_md_output/`. With no arguments the output looks like this:

```
Found 6 processable files across 5 extensions.

Processing 1 files for extension: '.txt'...
Processing 1 files for extension: '.pptx'...
Processing 1 files for extension: '.docx'...
Processing 1 files for extension: '.csv'...
Processing 2 files for extension: '.xlsx'...

All done! Converted files are saved in: /path/to/merged_md_output
```

Both folders are relative to the directory you run the command from, not to the script's location. Files that cannot be parsed are reported as `[Skipped] <name>: <error>` in the console, so check the output after a run.

### Converting specific files

Each argument is looked up in this order, and the first match wins:

1. The path as you typed it, relative to the current directory (so files outside `resources/` work too).
2. The path under `resources/`, for example `sub/data.xlsx`.
3. If the argument is a bare file name, a recursive search of `resources/` for a file with that name. If several files share the name, the command stops and lists them; give the path from `resources/` instead (`a/dup.txt`).

Naming the same file more than once converts it once. If any argument cannot be found or is not a supported type, the command reports every problem and exits without writing anything.

## Output format

**All files (no arguments):** output is grouped by extension and named `merged_<ext>_partNN.md`. Each file holds at most 100 source files. Bigger sets roll over to `part02`, `part03` and so on, which keeps each file at a size that is easy to upload or load into a context window.

```
merged_md_output/
├── merged_csv_part01.md
├── merged_docx_part01.md
├── merged_pptx_part01.md
├── merged_txt_part01.md
└── merged_xlsx_part01.md
```

**Specific files:** everything goes into a single file, whatever the file types, and it is not split into parts. One file is written as `<name>.md` (`budget.docx` becomes `budget.md`); two or more are merged into `merged.md`, in the order you gave them. These names never overlap with the `merged_<ext>_partNN.md` files, so converting a few files does not overwrite the results of a full run. It does overwrite an earlier output with the same name.

Every source file starts with a banner showing its path relative to `resources/` (or the path you typed, for a file outside `resources/`), so the AI can tell where each piece of content came from:

```markdown
================================================================================
# FILE: sub/data.xlsx
================================================================================

## Sheet: Sheet1

|   a | b   |
|----:|:----|
|   1 | x   |
|   2 | y   |
```

### Conversion details

- **Word:** every heading level becomes `###`. Tables are converted to Markdown tables, with newlines inside cells replaced by spaces.
- **Excel / CSV:** each sheet becomes a `## Sheet: <name>` section. Fully empty rows, fully empty columns and empty `Unnamed:` columns are dropped. Sheets with more than 15 columns are written as one `### Row N` block per row with `- **column**: value` lines (empty cells omitted), because a 100-column Markdown table is unreadable for both people and models.
- **PowerPoint:** each slide becomes a `### Slide N` section containing the text of its shapes. If `python-pptx` cannot open the file (for example some `.potx` templates), the script falls back to reading the slide XML directly.
- **PDF:** each page with extractable text becomes a `### Page N` section.
- **Text:** read as UTF-8; undecodable bytes are ignored.

## Configuration

The only command-line arguments are the optional files to convert. The two folders are set at the bottom of [convert_to_markdown.py](convert_to_markdown.py):

```python
source_dir = "resources"
output_dir = "./merged_md_output"
```

The 100-files-per-output limit is the `max_files_per_batch=100` parameter of `process_all_extensions()`.

You can also import both modes from your own code:

```python
from convert_to_markdown import convert_files, process_all_extensions

process_all_extensions("my_docs", "out", max_files_per_batch=50)
convert_files(["budget.docx", "sales.xlsx"], "my_docs", "out")
```

`convert_files()` raises `ValueError` if a file is missing, ambiguous or unsupported, or if none of the files could be converted.

The wide-table threshold is the `max_cols_for_table=15` default of `extract_excel()`.

### Adding a file type

Write a function that takes a `pathlib.Path` and returns a Markdown string, then register it in the module-level `EXTRACTORS` dict, which both modes use:

```python
def extract_md(file_path):
    return file_path.read_text(encoding="utf-8")

EXTRACTORS = {
    ...,
    ".md": extract_md,
}
```

## Known limitations

- **Legacy `.doc`:** requires LibreOffice. If the conversion fails, the output contains a `[Warning: Could not parse legacy .doc file ...]` placeholder instead of the content. An empty `_temp_doc_conv/` folder is left next to the source file.
- **Word:** tables are appended after all paragraphs of the document, not at their original position. Images, headers and footers are ignored.
- **PowerPoint:** text inside tables and speaker notes are not extracted. Template (`.potx`) files usually contain no slides, so they produce `[Template file contains no text slides]`.
- **PDF:** only the embedded text layer is read. Scanned or image-only PDFs come out empty because there is no OCR.
- **Images:** ignored entirely. Embedded images are not described or OCR'd.
- **CSV:** read with pandas defaults (comma-separated, UTF-8).
- **Unsupported types** such as `.xls`, `.ppt` and `.md` are skipped without a message when scanning `resources/`. If you name one on the command line, it is an error instead.
