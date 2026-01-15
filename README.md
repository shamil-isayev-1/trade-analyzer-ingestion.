# Trade Analyzer: Data Ingestion Layer

This project is a specialized tool designed to automate the ingestion and normalization of trading reports for further financial analysis.

## Key Technical Features

* **Data Ingestion & Normalization**: Built a pipeline to handle locale-specific decimal separators (comma to dot conversion) and enforced numeric types to ensure calculation determinism.
* **BOM Handling**: Utilizes `utf-8-sig` encoding to automatically strip Byte Order Mark characters, preventing indexing errors during CSV parsing.
* **Automated Formatting**: Integrated `openpyxl` for visual post-processing, including dynamic column width calculation and cell alignment for immediate readability in Excel.

## Tech Stack
* **Python 3.x**
* **Pandas**: For data manipulation and cleaning.
* **Openpyxl**: For advanced Excel formatting and styling.

## How to Use
1. Place your broker's CSV report in the project directory.
2. Update the `INPUT_PATH` variable in the script.
3. Run the script to generate a formatted `Trade_Report.xlsx`.
