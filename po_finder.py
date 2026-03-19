import os
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import xml.etree.ElementTree as ET
import pandas as pd
from collections import defaultdict
from typing import List, Optional, Tuple, Set

# ── Namespaces common in Mexican CFDI XMLs ──────────────────────────────────
CFDI_NS = {
    "cfdi": "http://www.sat.gob.mx/cfd/4",
    "cfdi3": "http://www.sat.gob.mx/cfd/3",
}

# ── Colour palette ────────────────────────────────────────────────────────────
BG       = "#1e1e2e"
CARD     = "#2a2a3e"
ACCENT   = "#7c6af7"
ACCENT2  = "#5dd6b5"
TEXT     = "#e0e0f0"
SUBTEXT  = "#8888aa"
RED      = "#f28b82"
BTN_FG   = "#ffffff"


def read_po_numbers(filepath: str, column: str) -> List[str]:
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xls", ".xlsx"):
        df = pd.read_excel(filepath, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(filepath, dtype=str)
    else:  # .txt  – try tab then comma
        try:
            df = pd.read_csv(filepath, sep="\t", dtype=str)
        except Exception:
            df = pd.read_csv(filepath, dtype=str)
    return df[column].dropna().str.strip().unique().tolist()


def search_xml_for_po(xml_path: str, po_numbers: Set[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Return (po_found, rfc, empresa) or (None, None, None)."""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        text_content = ET.tostring(root, encoding="unicode")

        matched_po = None
        for po in po_numbers:
            if po in text_content:
                matched_po = po
                break

        if matched_po is None:
            return None, None, None

        # Extract RFC + Nombre from CFDI (v3 or v4)
        rfc, empresa = None, None
        for prefix, ns in [("cfdi", CFDI_NS["cfdi"]), ("cfdi3", CFDI_NS["cfdi3"])]:
            emisor = root.find(f"{{{ns}}}Emisor")
            if emisor is None:
                emisor = root.find(f".//{{{ns}}}Emisor")
            if emisor is not None:
                rfc    = emisor.get("Rfc") or emisor.get("rfc")
                empresa = emisor.get("Nombre") or emisor.get("nombre")
                break

        # Fallback: search attributes root-level
        if rfc is None:
            rfc    = root.get("RfcEmisor") or root.get("RfcProveedor")
            empresa = root.get("NombreEmisor") or root.get("NombreProveedor")

        return matched_po, rfc, empresa

    except Exception:
        return None, None, None


# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PO → XML Finder  •  by Request")
        self.geometry("780x680")
        self.configure(bg=BG)
        self.resizable(True, True)

        # State
        self.data_file   = tk.StringVar()
        self.xml_dir     = tk.StringVar()
        self.dest_dir    = tk.StringVar()
        self.po_column   = tk.StringVar()
        self.copy_mode   = tk.StringVar(value="flat")   # "flat" or "structure"
        self.available_columns: List[str] = []

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        self._header()
        frame = tk.Frame(self, bg=BG, padx=24, pady=8)
        frame.pack(fill="both", expand=True)

        self._file_row(frame, "📄  Data file (XLSX / XLS / CSV / TXT)",
                       self.data_file, self._browse_data, row=0)
        self._column_selector(frame, row=1)
        self._file_row(frame, "📂  XML source folder",
                       self.xml_dir, self._browse_xml_dir, row=2, is_dir=True)
        self._file_row(frame, "📁  Destination folder",
                       self.dest_dir, self._browse_dest_dir, row=3, is_dir=True)

        self._copy_mode_selector(frame, row=4)
        self._run_button(frame)
        self._log_area(frame)

    def _header(self):
        hdr = tk.Frame(self, bg=ACCENT, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🔎  PO → XML Finder", font=("Segoe UI", 18, "bold"),
                 bg=ACCENT, fg=BTN_FG).pack()
        tk.Label(hdr, text="Locate, match and copy XML files by Purchase Order number",
                 font=("Segoe UI", 10), bg=ACCENT, fg="#ddddff").pack()

    def _file_row(self, parent, label, var, cmd, row, is_dir=False):
        tk.Label(parent, text=label, bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=row*3,   column=0, columnspan=3,
                                             sticky="w", pady=(12, 0))
        entry = tk.Entry(parent, textvariable=var, bg=CARD, fg=TEXT,
                         insertbackground=TEXT, relief="flat",
                         font=("Segoe UI", 10), bd=6)
        entry.grid(row=row*3+1, column=0, columnspan=2, sticky="ew", ipady=4)
        btn = tk.Button(parent, text="Browse", command=cmd,
                        bg=ACCENT, fg=BTN_FG, relief="flat",
                        font=("Segoe UI", 9, "bold"), padx=12, cursor="hand2")
        btn.grid(row=row*3+1, column=2, sticky="ew", padx=(6, 0))
        parent.columnconfigure(0, weight=1)

    def _column_selector(self, parent, row):
        tk.Label(parent, text="📌  PO Number column", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=row*3, column=0, columnspan=3,
                                             sticky="w", pady=(12, 0))
        inner = tk.Frame(parent, bg=BG)
        inner.grid(row=row*3+1, column=0, columnspan=3, sticky="ew")
        inner.columnconfigure(0, weight=1)

        self.col_combo = ttk.Combobox(inner, textvariable=self.po_column,
                                      state="readonly", font=("Segoe UI", 10))
        self.col_combo.grid(row=0, column=0, sticky="ew", ipady=3)

        style = ttk.Style()
        style.configure("TCombobox", fieldbackground=CARD, background=CARD,
                        foreground=TEXT)

        tk.Button(inner, text="Auto-detect", command=self._auto_detect_column,
                  bg=ACCENT2, fg="#000000", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=10, cursor="hand2"
                  ).grid(row=0, column=1, padx=(6, 0))

    def _copy_mode_selector(self, parent, row):
        tk.Label(parent, text="📋  Destination structure", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=row*3, column=0, columnspan=3,
                                              sticky="w", pady=(12, 0))
        inner = tk.Frame(parent, bg=BG)
        inner.grid(row=row*3+1, column=0, columnspan=3, sticky="w")
        tk.Radiobutton(inner, text="Flat — all XMLs in one folder",
                       variable=self.copy_mode, value="flat",
                       bg=BG, fg=TEXT, selectcolor=CARD,
                       activebackground=BG, activeforeground=ACCENT2,
                       font=("Segoe UI", 10)).pack(side="left", padx=(0, 20))
        tk.Radiobutton(inner, text="Keep folder structure",
                       variable=self.copy_mode, value="structure",
                       bg=BG, fg=TEXT, selectcolor=CARD,
                       activebackground=BG, activeforeground=ACCENT2,
                       font=("Segoe UI", 10)).pack(side="left")

    def _run_button(self, parent):
        tk.Button(parent, text="▶  Run — Find & Copy XMLs",
                  command=self._run_threaded,
                  bg=ACCENT, fg=BTN_FG, relief="flat",
                  font=("Segoe UI", 12, "bold"), pady=10, cursor="hand2"
                  ).grid(row=14, column=0, columnspan=3, sticky="ew", pady=(20, 6))

    def _log_area(self, parent):
        tk.Label(parent, text="📋  Log", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(row=15, column=0, sticky="w")

        self.progress = ttk.Progressbar(parent, mode="determinate")
        self.progress.grid(row=16, column=0, columnspan=3, sticky="ew", pady=(0, 6))

        self.log = tk.Text(parent, bg=CARD, fg=TEXT, relief="flat",
                           font=("Consolas", 9), height=12, wrap="word",
                           insertbackground=TEXT, state="disabled")
        self.log.grid(row=17, column=0, columnspan=3, sticky="nsew")
        sb = tk.Scrollbar(parent, command=self.log.yview, bg=CARD)
        sb.grid(row=17, column=3, sticky="ns")
        self.log["yscrollcommand"] = sb.set
        parent.rowconfigure(17, weight=1)

    # ── Browse callbacks ──────────────────────────────────────────────────────
    def _browse_data(self):
        path = filedialog.askopenfilename(
            title="Select data file",
            filetypes=[("Spreadsheet / Text", "*.xlsx *.xls *.csv *.txt"),
                       ("All files", "*.*")])
        if path:
            self.data_file.set(path)
            self._load_columns(path)

    def _browse_xml_dir(self):
        d = filedialog.askdirectory(title="Select folder containing XML files")
        if d:
            self.xml_dir.set(d)

    def _browse_dest_dir(self):
        d = filedialog.askdirectory(title="Select destination folder")
        if d:
            self.dest_dir.set(d)

    # ── Column helpers ────────────────────────────────────────────────────────
    def _load_columns(self, path: str):
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in (".xls", ".xlsx"):
                df = pd.read_excel(path, nrows=0)
            elif ext == ".csv":
                df = pd.read_csv(path, nrows=0)
            else:
                try:    df = pd.read_csv(path, sep="\t", nrows=0)
                except: df = pd.read_csv(path, nrows=0)
            self.available_columns = df.columns.tolist()
            self.col_combo["values"] = self.available_columns
            # pre-select if obvious PO column found
            for c in self.available_columns:
                if "po" in c.lower() or "order" in c.lower() or "purchase" in c.lower():
                    self.po_column.set(c)
                    break
            else:
                if self.available_columns:
                    self.po_column.set(self.available_columns[0])
        except Exception as e:
            self._log(f"[WARN] Could not read columns: {e}", RED)

    def _auto_detect_column(self):
        if not self.data_file.get():
            messagebox.showwarning("No file", "Please select a data file first.")
            return
        self._load_columns(self.data_file.get())
        self._log(f"Columns loaded: {self.available_columns}", ACCENT2)
        messagebox.showinfo("Columns detected",
                            f"Found {len(self.available_columns)} columns.\n"
                            f"Selected: «{self.po_column.get()}»\n\n"
                            "You can change it in the dropdown.")

    # ── Main logic (threaded) ─────────────────────────────────────────────────
    def _run_threaded(self):
        t = threading.Thread(target=self._run, daemon=True)
        t.start()

    def _run(self):
        # Validate inputs
        if not self.data_file.get():
            return messagebox.showerror("Missing", "Please select a data file.")
        if not self.po_column.get():
            return messagebox.showerror("Missing", "Please select the PO column.")
        if not self.xml_dir.get():
            return messagebox.showerror("Missing", "Please select the XML source folder.")
        if not self.dest_dir.get():
            return messagebox.showerror("Missing", "Please select the destination folder.")

        self._log("─" * 60)
        self._log("▶  Starting process…")

        # 1. Read PO numbers
        try:
            po_list = read_po_numbers(self.data_file.get(), self.po_column.get())
            po_set  = set(po_list)
            self._log(f"✔  PO numbers loaded: {len(po_set)}", ACCENT2)
        except Exception as e:
            self._log(f"✘  Error reading data file: {e}", RED)
            return

        # 2. Scan XML files recursively through all subfolders
        xml_files = []
        for dirpath, _, filenames in os.walk(self.xml_dir.get()):
            for f in filenames:
                if f.lower().endswith(".xml"):
                    xml_files.append(os.path.join(dirpath, f))

        if not xml_files:
            self._log("✘  No XML files found in selected folder or subfolders.", RED)
            return
        self._log(f"✔  XML files found (all subfolders): {len(xml_files)}", ACCENT2)

        os.makedirs(self.dest_dir.get(), exist_ok=True)

        self.progress["maximum"] = len(xml_files)
        self.progress["value"]   = 0

        copied        = 0
        not_found_pos = set(po_set.copy())
        rfc_empresa   = defaultdict(set)
        po_to_files   = defaultdict(list)
        xml_root      = self.xml_dir.get()

        for i, full_path in enumerate(xml_files):
            fname = os.path.basename(full_path)
            matched_po, rfc, empresa = search_xml_for_po(full_path, po_set)

            if matched_po:
                if self.copy_mode.get() == "structure":
                    # Recreate relative subfolder path inside destination
                    rel_path  = os.path.relpath(full_path, xml_root)
                    dest_path = os.path.join(self.dest_dir.get(), rel_path)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                else:
                    dest_path = os.path.join(self.dest_dir.get(), fname)

                shutil.copy2(full_path, dest_path)
                copied += 1
                not_found_pos.discard(matched_po)
                po_to_files[matched_po].append(fname)
                if rfc:
                    rfc_empresa[rfc].add(empresa or "—")
                rel_display = os.path.relpath(full_path, xml_root)
                self._log(f"  ✔  {rel_display}  →  PO: {matched_po}", ACCENT2)

            self.progress["value"] = i + 1
            self.update_idletasks()

        # ── Summary ──────────────────────────────────────────────────────────
        self._log("─" * 60)
        self._log(f"✅  Process complete!")
        self._log(f"    Files copied  : {copied}")
        self._log(f"    POs matched   : {len(po_set) - len(not_found_pos)} / {len(po_set)}")

        if not_found_pos:
            self._log(f"    POs not found : {', '.join(sorted(not_found_pos))}", RED)

        rfc_summary_lines = []
        for rfc, names in rfc_empresa.items():
            line = f"RFC: {rfc}  |  Empresa: {', '.join(names)}"
            rfc_summary_lines.append(line)
            self._log(f"    {line}", ACCENT2)

        # ── Pop-up summary ────────────────────────────────────────────────────
        rfc_block = "\n".join(rfc_summary_lines) if rfc_summary_lines else "No RFC info found in XMLs"
        not_found_block = (
            f"\n\n⚠️  POs not found ({len(not_found_pos)}):\n" +
            "\n".join(sorted(not_found_pos))
        ) if not_found_pos else ""

        messagebox.showinfo(
            "✅  Process Complete",
            f"Files copied: {copied}\n"
            f"POs matched: {len(po_set) - len(not_found_pos)} / {len(po_set)}\n\n"
            f"── Suppliers / Proveedores ──\n{rfc_block}"
            f"{not_found_block}"
        )

    # ── Log helper ────────────────────────────────────────────────────────────
    def _log(self, msg: str, color: str = TEXT):
        def _do():
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n", color)
            self.log.tag_configure(color, foreground=color)
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _do)


if __name__ == "__main__":
    app = App()
    app.mainloop()