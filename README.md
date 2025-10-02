# converter_xlsxtofsh

Excel → FSH Converter
======================

This script converts large (10k+ rows) Excel/CSV/TSV files with medication codes and names into **FHIR Shorthand (FSH)** format.

Output structure example:

```
* #<code> "<uz>"
  * ^designation[0].language = #ru
  * ^designation[=].value = "<russian>"
  * ^designation[+].language = #en
  * ^designation[=].value = "<english>"
  * ^designation[+].language = #la
  * ^designation[=].value = "<latin>"
```

Additional languages (columns like `lang:xx`) are automatically appended at the end.

---

Installation
------------

1. Clone the repository or download the script:
2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1   # Windows
   source .venv/bin/activate        # Linux/Mac
   ```
3. Install dependencies:
   ```bash
   pip install pandas openpyxl
   ```

---

Usage
-----

### For Excel files
```bash
python tools/xlsx_to_fsh_uz_ru_en_la.py data/input.xlsx -o out/meds.fsh --sheet Sheet1 --code CodeCol --uz UzbekCol --ru RussianCol --en EnglishCol --la LatinCol
```

### For CSV files
```bash
python tools/xlsx_to_fsh_uz_ru_en_la.py data/input.csv -o out/meds.fsh --code code --uz uz --ru ru --en en --la la
```

---

Parameters
----------

- `--code` – column for concept codes (required)  
- `--uz` – column for Uzbek display (required, main display value)  
- `--ru` – column for Russian designation  
- `--en` – column for English designation  
- `--la` – column for Latin designation  
- `--sheet` – Excel sheet name (optional)  
- `-o` – output file path (default: `<input>.fsh`)  

Additional languages: any column named `lang:xx` (e.g., `lang:kk`) will be added automatically.

---

Output
------

The resulting `.fsh` file will contain one entry per row. The main display is Uzbek, and designations follow the order: **ru → en → la → (extras)**.

---

Example
-------

Excel input:

| Code | UzbekCol | RussianCol | EnglishCol | LatinCol |
|------|----------|------------|------------|----------|
| 0001 | Abay     | Абай       | Abay       | Abaium   |

Generated FSH:

```
* #0001 "Abay"
  * ^designation[0].language = #ru
  * ^designation[=].value = "Абай"
  * ^designation[+].language = #en
  * ^designation[=].value = "Abay"
  * ^designation[+].language = #la
  * ^designation[=].value = "Abaium"
