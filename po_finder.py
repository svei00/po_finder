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
    "cfdi":  "http://www.sat.gob.mx/cfd/4",
    "cfdi3": "http://www.sat.gob.mx/cfd/3",
}

# ── Colour palette ────────────────────────────────────────────────────────────
BG      = "#1e1e2e"
CARD    = "#2a2a3e"
ACCENT  = "#7c6af7"
ACCENT2 = "#5dd6b5"
YELLOW  = "#f9e04b"
TEXT    = "#e0e0f0"
SUBTEXT = "#8888aa"
RED     = "#f28b82"
BTN_FG  = "#ffffff"


# ─────────────────────────────────────────────────────────────────────────────
def read_po_numbers(filepath, column):
    ext = os.path.splitext(filepath)[1].lower()
    if ext in (".xls", ".xlsx"):
        df = pd.read_excel(filepath, dtype=str)
    elif ext == ".csv":
        df = pd.read_csv(filepath, dtype=str)
    else:
        try:
            df = pd.read_csv(filepath, sep="\t", dtype=str)
        except Exception:
            df = pd.read_csv(filepath, dtype=str)
    return df[column].dropna().str.strip().unique().tolist()


def _extract_party(root, tag):
    """Extract (Rfc, Nombre) for Emisor or Receptor node."""
    for ns in [CFDI_NS["cfdi"], CFDI_NS["cfdi3"]]:
        node = root.find(f"{{{ns}}}{tag}")
        if node is None:
            node = root.find(f".//{{{ns}}}{tag}")
        if node is not None:
            rfc    = node.get("Rfc")    or node.get("rfc")
            nombre = node.get("Nombre") or node.get("nombre")
            return rfc, nombre
    return None, None


def search_xml_for_po(xml_path, po_numbers):
    """
    Returns (matched_pos, emisor_rfc, emisor_nombre, receptor_rfc, receptor_nombre).
    matched_pos is a LIST of all POs found inside this XML.
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        text_content = ET.tostring(root, encoding="unicode")

        # FIX: collect ALL matching POs, not just the first
        matched_pos = [po for po in po_numbers if po in text_content]
        if not matched_pos:
            return [], None, None, None, None

        emisor_rfc,   emisor_nombre   = _extract_party(root, "Emisor")
        receptor_rfc, receptor_nombre = _extract_party(root, "Receptor")

        if emisor_rfc is None:
            emisor_rfc    = root.get("RfcEmisor")
            emisor_nombre = root.get("NombreEmisor")
        if receptor_rfc is None:
            receptor_rfc    = root.get("RfcReceptor")
            receptor_nombre = root.get("NombreReceptor")

        return matched_pos, emisor_rfc, emisor_nombre, receptor_rfc, receptor_nombre

    except Exception:
        return [], None, None, None, None


# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PO -> XML Finder  |  CFDI Mexico")
        self.geometry("820x720")
        self.configure(bg=BG)
        self.resizable(True, True)

        self.data_file         = tk.StringVar()
        self.xml_dir           = tk.StringVar()
        self.dest_dir          = tk.StringVar()
        self.po_column         = tk.StringVar()
        self.copy_mode         = tk.StringVar(value="flat")
        self.available_columns = []

        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._header()
        frame = tk.Frame(self, bg=BG, padx=24, pady=8)
        frame.pack(fill="both", expand=True)

        self._file_row(frame, "Data file (XLSX / XLS / CSV / TXT)",
                       self.data_file, self._browse_data, row=0)
        self._column_selector(frame, row=1)
        self._file_row(frame, "XML source folder  (subfolders scanned automatically)",
                       self.xml_dir, self._browse_xml_dir, row=2)
        self._file_row(frame, "Destination folder",
                       self.dest_dir, self._browse_dest_dir, row=3)
        self._copy_mode_selector(frame, row=4)
        self._run_button(frame)
        self._log_area(frame)

    def _header(self):
        hdr = tk.Frame(self, bg=ACCENT, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  PO -> XML Finder",
                 font=("Segoe UI", 18, "bold"), bg=ACCENT, fg=BTN_FG).pack()
        tk.Label(hdr, text="Locate, match and copy CFDI XML files by Purchase Order number",
                 font=("Segoe UI", 10), bg=ACCENT, fg="#ddddff").pack()

    def _file_row(self, parent, label, var, cmd, row):
        tk.Label(parent, text=label, bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(
                     row=row*3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        tk.Entry(parent, textvariable=var, bg=CARD, fg=TEXT,
                 insertbackground=TEXT, relief="flat",
                 font=("Segoe UI", 10), bd=6).grid(
                     row=row*3+1, column=0, columnspan=2, sticky="ew", ipady=4)
        tk.Button(parent, text="Browse", command=cmd, bg=ACCENT, fg=BTN_FG,
                  relief="flat", font=("Segoe UI", 9, "bold"),
                  padx=12, cursor="hand2").grid(
                      row=row*3+1, column=2, sticky="ew", padx=(6, 0))
        parent.columnconfigure(0, weight=1)

    def _column_selector(self, parent, row):
        tk.Label(parent, text="PO Number column", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(
                     row=row*3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        inner = tk.Frame(parent, bg=BG)
        inner.grid(row=row*3+1, column=0, columnspan=3, sticky="ew")
        inner.columnconfigure(0, weight=1)

        self.col_combo = ttk.Combobox(inner, textvariable=self.po_column,
                                      state="readonly", font=("Segoe UI", 10))
        self.col_combo.grid(row=0, column=0, sticky="ew", ipady=3)
        ttk.Style().configure("TCombobox", fieldbackground=CARD,
                               background=CARD, foreground=TEXT)
        tk.Button(inner, text="Auto-detect", command=self._auto_detect_column,
                  bg=ACCENT2, fg="#000000", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=10, cursor="hand2"
                  ).grid(row=0, column=1, padx=(6, 0))

    def _copy_mode_selector(self, parent, row):
        tk.Label(parent, text="Destination structure", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).grid(
                     row=row*3, column=0, columnspan=3, sticky="w", pady=(12, 0))
        inner = tk.Frame(parent, bg=BG)
        inner.grid(row=row*3+1, column=0, columnspan=3, sticky="w")
        for text, val in [("Flat - all XMLs in one folder", "flat"),
                          ("Keep original folder structure",  "structure")]:
            tk.Radiobutton(inner, text=text, variable=self.copy_mode, value=val,
                           bg=BG, fg=TEXT, selectcolor=CARD,
                           activebackground=BG, activeforeground=ACCENT2,
                           font=("Segoe UI", 10)).pack(side="left", padx=(0, 24))

    def _run_button(self, parent):
        tk.Button(parent, text="Run - Find & Copy XMLs",
                  command=self._run_threaded, bg=ACCENT, fg=BTN_FG,
                  relief="flat", font=("Segoe UI", 12, "bold"),
                  pady=10, cursor="hand2").grid(
                      row=14, column=0, columnspan=3, sticky="ew", pady=(20, 6))

    def _log_area(self, parent):
        # Header row with Copy + Clear buttons
        hdr = tk.Frame(parent, bg=BG)
        hdr.grid(row=15, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        tk.Label(hdr, text="Log", bg=BG, fg=SUBTEXT,
                 font=("Segoe UI", 9)).pack(side="left")
        tk.Button(hdr, text="Copy Log", command=self._copy_log,
                  bg=CARD, fg=ACCENT2, relief="flat",
                  font=("Segoe UI", 8, "bold"), padx=8, cursor="hand2"
                  ).pack(side="right")
        tk.Button(hdr, text="Clear", command=self._clear_log,
                  bg=CARD, fg=SUBTEXT, relief="flat",
                  font=("Segoe UI", 8), padx=8, cursor="hand2"
                  ).pack(side="right", padx=(0, 4))

        self.progress = ttk.Progressbar(parent, mode="determinate")
        self.progress.grid(row=16, column=0, columnspan=3,
                           sticky="ew", pady=(4, 6))

        # state="normal" so text is selectable; typing blocked via binding
        self.log = tk.Text(parent, bg=CARD, fg=TEXT, relief="flat",
                           font=("Consolas", 9), height=12, wrap="word",
                           insertbackground=TEXT, cursor="arrow")
        self.log.grid(row=17, column=0, columnspan=3, sticky="nsew")
        sb = tk.Scrollbar(parent, command=self.log.yview, bg=CARD)
        sb.grid(row=17, column=3, sticky="ns")
        self.log["yscrollcommand"] = sb.set
        parent.rowconfigure(17, weight=1)
        self.log.bind("<Key>", self._block_typing)

    # ── Browse ────────────────────────────────────────────────────────────────
    def _browse_data(self):
        path = filedialog.askopenfilename(
            title="Select data file",
            filetypes=[("Spreadsheet / Text", "*.xlsx *.xls *.csv *.txt"),
                       ("All files", "*.*")])
        if path:
            self.data_file.set(path)
            self._load_columns(path)

    def _browse_xml_dir(self):
        d = filedialog.askdirectory(title="Select XML source folder")
        if d:
            self.xml_dir.set(d)

    def _browse_dest_dir(self):
        d = filedialog.askdirectory(title="Select destination folder")
        if d:
            self.dest_dir.set(d)

    # ── Column helpers ────────────────────────────────────────────────────────
    def _load_columns(self, path):
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
                            f"Selected: [{self.po_column.get()}]\n\n"
                            "You can change it in the dropdown.")

    # ── Log helpers ───────────────────────────────────────────────────────────
    def _log(self, msg, color=None):
        if color is None:
            color = TEXT
        def _do():
            self.log.insert("end", msg + "\n", color)
            self.log.tag_configure(color, foreground=color)
            self.log.see("end")
        self.after(0, _do)

    def _copy_log(self):
        content = self.log.get("1.0", "end")
        self.clipboard_clear()
        self.clipboard_append(content)
        messagebox.showinfo("Copied", "Log copied to clipboard.")

    def _clear_log(self):
        self.log.delete("1.0", "end")
        self.progress["value"] = 0

    def _block_typing(self, event):
        allowed_keys = (
            "Up", "Down", "Left", "Right", "Home", "End",
            "Prior", "Next", "Shift_L", "Shift_R",
            "Control_L", "Control_R", "Alt_L", "Alt_R"
        )
        if event.keysym in allowed_keys:
            return
        # Allow Ctrl+C and Ctrl+A
        if event.state & 0x4 and event.keysym.lower() in ("c", "a"):
            return
        return "break"

    # ── Main logic ────────────────────────────────────────────────────────────
    def _run_threaded(self):
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        if not self.data_file.get():
            return messagebox.showerror("Missing", "Please select a data file.")
        if not self.po_column.get():
            return messagebox.showerror("Missing", "Please select the PO column.")
        if not self.xml_dir.get():
            return messagebox.showerror("Missing", "Please select the XML source folder.")
        if not self.dest_dir.get():
            return messagebox.showerror("Missing", "Please select the destination folder.")

        self._log("=" * 60)
        self._log("Starting process...")

        # 1. Read PO numbers
        try:
            po_list = read_po_numbers(self.data_file.get(), self.po_column.get())
            po_set  = set(po_list)
            self._log(f"OK  PO numbers loaded: {len(po_set)}", ACCENT2)
        except Exception as e:
            self._log(f"ERROR reading data file: {e}", RED)
            return

        # 2. Walk all subfolders
        xml_files = []
        for dirpath, _, filenames in os.walk(self.xml_dir.get()):
            for f in filenames:
                if f.lower().endswith(".xml"):
                    xml_files.append(os.path.join(dirpath, f))

        if not xml_files:
            self._log("ERROR: No XML files found in folder or subfolders.", RED)
            return
        self._log(f"OK  XML files found (all subfolders): {len(xml_files)}", ACCENT2)

        os.makedirs(self.dest_dir.get(), exist_ok=True)
        self.progress["maximum"] = len(xml_files)
        self.progress["value"]   = 0

        copied        = 0
        skipped       = 0
        not_found_pos = set(po_set)
        emisor_map    = defaultdict(set)
        receptor_map  = defaultdict(set)
        xml_root      = self.xml_dir.get()

        for i, full_path in enumerate(xml_files):
            fname = os.path.basename(full_path)
            matched_pos, e_rfc, e_name, r_rfc, r_name = search_xml_for_po(
                full_path, po_set)

            if matched_pos:
                # Determine destination
                if self.copy_mode.get() == "structure":
                    rel       = os.path.relpath(full_path, xml_root)
                    dest_path = os.path.join(self.dest_dir.get(), rel)
                    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
                else:
                    dest_path = os.path.join(self.dest_dir.get(), fname)

                rel_display = os.path.relpath(full_path, xml_root)
                po_display  = ", ".join(sorted(matched_pos))

                # FIX 4: skip if already exists
                if os.path.exists(dest_path):
                    skipped += 1
                    self._log(f"  SKIP  {rel_display}  (already in destination)", YELLOW)
                else:
                    shutil.copy2(full_path, dest_path)
                    copied += 1
                    self._log(f"  COPY  {rel_display}  |  PO(s): {po_display}", ACCENT2)

                # Mark ALL matched POs as found
                for po in matched_pos:
                    not_found_pos.discard(po)

                if e_rfc:
                    emisor_map[e_rfc].add(e_name or "---")
                if r_rfc:
                    receptor_map[r_rfc].add(r_name or "---")

            self.progress["value"] = i + 1
            self.update_idletasks()

        # ── Summary ──────────────────────────────────────────────────────────
        self._log("=" * 60)
        self._log("PROCESS COMPLETE")
        self._log(f"  Files copied  : {copied}")
        self._log(f"  Files skipped : {skipped}  (already existed in destination)")
        self._log(f"  POs matched   : {len(po_set) - len(not_found_pos)} / {len(po_set)}")

        if not_found_pos:
            self._log(
                f"  POs NOT found : {', '.join(sorted(not_found_pos))}", RED)

        emisor_lines   = []
        receptor_lines = []

        self._log("  -- Emisor / Proveedor --", ACCENT2)
        for rfc, names in emisor_map.items():
            line = f"    RFC: {rfc}  |  Nombre: {', '.join(sorted(names))}"
            emisor_lines.append(line)
            self._log(line, ACCENT2)

        self._log("  -- Receptor / Cliente --", ACCENT2)
        for rfc, names in receptor_map.items():
            line = f"    RFC: {rfc}  |  Nombre: {', '.join(sorted(names))}"
            receptor_lines.append(line)
            self._log(line, ACCENT2)

        # ── Popup ─────────────────────────────────────────────────────────────
        e_block = "\n".join(emisor_lines)   or "  Not found in XMLs"
        r_block = "\n".join(receptor_lines) or "  Not found in XMLs"
        nf_block = (
            f"\n\nPOs NOT found ({len(not_found_pos)}):\n"
            + "\n".join(sorted(not_found_pos))
        ) if not_found_pos else ""

        messagebox.showinfo(
            "Process Complete",
            f"Files copied : {copied}\n"
            f"Files skipped: {skipped}  (already existed)\n"
            f"POs matched  : {len(po_set) - len(not_found_pos)} / {len(po_set)}\n"
            f"\n-- Emisor / Proveedor --\n{e_block}"
            f"\n\n-- Receptor / Cliente --\n{r_block}"
            f"{nf_block}"
        )


if __name__ == "__main__":
    app = App()
    app.mainloop()