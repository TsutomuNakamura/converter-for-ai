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

1. Put the files you want to convert in a `resources/` folder. Subfolders are searched recursively.
2. Run the script from the directory that contains `resources/`:

   ```bash
   python convert_to_markdown.py
   ```

3. Read the results in `merged_md_output/`.

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

## Output format

Output is grouped by extension and named `merged_<ext>_partNN.md`. Each file holds at most 100 source files. Bigger sets roll over to `part02`, `part03` and so on, which keeps each file at a size that is easy to upload or load into a context window.

```
merged_md_output/
├── merged_csv_part01.md
├── merged_docx_part01.md
├── merged_pptx_part01.md
├── merged_txt_part01.md
└── merged_xlsx_part01.md
```

Every source file starts with a banner showing its path relative to `resources/`, so the AI can tell where each piece of content came from:

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

The script has no command-line options. Edit the call at the bottom of [convert_to_markdown.py](convert_to_markdown.py):

```python
process_all_extensions(
    source_dir="resources",
    output_dir="./merged_md_output",
    max_files_per_batch=100,
)
```

You can also import it from your own code:

```python
from convert_to_markdown import process_all_extensions

process_all_extensions("my_docs", "out", max_files_per_batch=50)
```

The wide-table threshold is the `max_cols_for_table=15` default of `extract_excel()`.

### Adding a file type

Write a function that takes a `pathlib.Path` and returns a Markdown string, then register it in the `extractors` dict inside `process_all_extensions()`:

```python
def extract_md(file_path):
    return file_path.read_text(encoding="utf-8")

extractors = {
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
- **Unsupported types** such as `.xls`, `.ppt` and `.md` are skipped without a message.
