# 🔎 PO XML Finder

A desktop tool that reads Purchase Order numbers from a spreadsheet, searches a folder of CFDI XML files for matching POs, and copies the matching XMLs to a destination folder — with a full summary of RFC and supplier names.

---

## 📋 Features

- Supports **XLSX, XLS, CSV and TXT** input files
- **Dynamic column selector** — auto-detects the PO column or let you pick from a dropdown
- Searches XML files using **full content scan** (works regardless of which node the PO lives in)
- Compatible with **CFDI v3 and v4** (SAT Mexico format) for RFC and Empresa extraction
- Copies matched XML files to a destination folder of your choice
- Shows a **summary popup** with:
  - Total files copied
  - POs matched vs. not found
  - RFC and Empresa name for each supplier found
- Live log with progress bar during processing
- **No command line needed** — double-click launcher for Windows and Mac/Linux

---

## 🗂 File Structure

```
po_finder/
├── po_xml_finder.py      # Main application
├── run_po_finder.bat     # Windows launcher (double-click)
├── run_po_finder.sh      # Mac / Linux launcher (double-click)
└── README.md
```

---

## ⚙️ Requirements

- **Python 3.8+** — download from [python.org](https://www.python.org/downloads/)
  - Windows: during installation check ✅ **"Add Python to PATH"**
- The launchers auto-install required libraries on first run:
  - `pandas`
  - `openpyxl`

---

## 🚀 How to Run

### Windows
Double-click `run_po_finder.bat`

### Mac / Linux
First time only — open a terminal in the folder and run:
```bash
chmod +x run_po_finder.sh
```
Then double-click `run_po_finder.sh`

### Manual (any OS)
```bash
pip install pandas openpyxl
python po_xml_finder.py
```

---

## 🖥 How to Use

1. **Data file** — Browse and select your XLSX / XLS / CSV / TXT file containing PO numbers
2. **PO column** — The app auto-detects the PO column; use the dropdown or click **Auto-detect** to change it
3. **XML source folder** — Browse to the folder that contains your XML (CFDI) files
4. **Destination folder** — Browse to the folder where matched XMLs should be copied
5. Click **▶ Run** and watch the live log
6. A summary popup appears when done

---

## 📄 Input File Format

Your spreadsheet must have a column with PO numbers. Example:

| PO Number     | Requester       | Supplier         | Total      |
|---------------|-----------------|------------------|------------|
| C-10003437497 | Gustavo Navarro | Hugo González... | 394,383.50 |
| C-10003428642 | Karla Ramos     | Hugo González... | 121,460.00 |

The column name does not need to be "PO Number" — you can select any column via the dropdown.

---

## 🔍 How XML Matching Works

The tool searches the **full text content** of each XML file for each PO number string (e.g. `C-10003437497`). This means it finds the PO regardless of which XML node or attribute it appears in.

RFC and Empresa name are extracted from the `<cfdi:Emisor>` node (CFDI v3 and v4 both supported).

---

## 🐛 Troubleshooting

| Problem | Solution |
|---|---|
| `Python not found` on Windows | Reinstall Python and check "Add to PATH" |
| `Permission denied` on Mac/Linux | Run `chmod +x run_po_finder.sh` once |
| No XML files found | Make sure the source folder contains `.xml` files |
| PO not matched | Check that the PO value in your spreadsheet exactly matches what is inside the XML |
| Missing `pandas` error | Run `pip install pandas openpyxl` manually |

---

## 📃 License

MIT — free to use and modify.
