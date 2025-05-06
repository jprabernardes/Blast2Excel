# Script to Extract BLAST Hits from a GenBank File

This Python script parses a `.gb` (GenBank) file containing annotated BLAST hits and exports the extracted data to an Excel spreadsheet.

## Requirements

- Python 3.6 or higher

## Installing dependencies

Install all required packages using:

```bash
pip install -r requirements.txt
```

## Usage

1. Place your input file (`orig_with_blast.gb`) in the same directory as the script.

2. The script automatically detects the directory it is in using:

```python
# Get the directory of the current script, put the file in the same directory as the script
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
file = os.path.join(script_dir, "orig_with_blast.gb")
excel_output = os.path.join(script_dir, "blast_hits.xlsx")
```

3. Run the script:

```bash
python parse_blast_hits.py
```

4. The output will be an Excel file named `blast_hits.xlsx` created in the same directory.

## Dependencies

- pandas
- biopython
- openpyxl
