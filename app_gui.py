"""
BGB Notice Generator — Desktop GUI (Tkinter)
Packaged with PyInstaller into a standalone .app / .exe for offline use.
"""

import os
import sys
import threading
import zipfile
import shutil
import tempfile
from datetime import date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------------------------------------------------------------------
# Ensure bundled modules (PyInstaller) are importable
# ---------------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    # Running inside PyInstaller bundle
    BASE_DIR = sys._MEIPASS
    sys.path.insert(0, BASE_DIR)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, BASE_DIR)

from excel_parser import parse_excel
from letter_generator import generate_extension_notice, generate_encashment_notice
from utils import validity_in_window, format_date


# ---------------------------------------------------------------------------
# Colours / fonts
# ---------------------------------------------------------------------------
BG        = "#F0F2F5"   # light grey page background
CARD      = "#FFFFFF"   # white card panels
PRIMARY   = "#1A3C6E"   # dark navy — header + button bg
BTN_FG    = "#FFFFFF"   # white text on dark buttons
LABEL_FG  = "#1A1A2E"   # near-black for all labels and radio text
MUTED_FG  = "#555577"   # secondary / hint text
ACCENT    = "#2E7D32"   # green for success button
DANGER    = "#C62828"   # red for errors
BORDER    = "#C8CDD6"   # entry/separator border colour

FONT_MAIN = ("Helvetica", 12)
FONT_HEAD = ("Helvetica", 18, "bold")
FONT_SUB  = ("Helvetica", 11)
FONT_BOLD = ("Helvetica", 11, "bold")
FONT_MONO = ("Courier", 10)


def _btn(parent, text, command, bg=PRIMARY, fg=BTN_FG, font=FONT_SUB, **kw):
    """
    Cross-platform button that always shows correct colours on macOS.
    Uses relief='raised' + overrelief='groove' instead of 'flat', which
    prevents macOS from overriding the bg/fg with system colours.
    """
    # kw may override padx/pady, so set defaults only if not already in kw
    kw.setdefault("padx", 10)
    kw.setdefault("pady", 6)
    return tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, font=font,
        activebackground=bg, activeforeground=fg,
        relief="raised", overrelief="groove",
        bd=0, cursor="hand2", **kw,
    )


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------
class BGBApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("BGB Notice Generator")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._center_window(780, 620)

        self._excel_path   = tk.StringVar()
        self._output_path  = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "BGB_Output"))
        self._letter_type  = tk.StringVar(value="both")
        self._letter_num   = tk.StringVar(value="800542/___/E8")
        self._letter_date  = tk.StringVar(value=date.today().strftime("%Y-%m-%d"))
        self._specific_row = tk.StringVar(value="")
        self._status_lines = []

        self._build_ui()

    # ------------------------------------------------------------------
    # UI Construction
    # ------------------------------------------------------------------
    def _build_ui(self):
        # ── Header ────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=PRIMARY, pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="BGB Notice Generator",
                 font=FONT_HEAD, bg=PRIMARY, fg=BTN_FG).pack()
        tk.Label(hdr, text="Bank Guarantee Bond — Extension & Encashment Notices",
                 font=FONT_SUB, bg=PRIMARY, fg="#B0C4DE").pack()

        # ── Main content ──────────────────────────────────────────────
        body = tk.Frame(self, bg=BG, padx=24, pady=16)
        body.pack(fill="both", expand=True)

        # Excel file picker
        self._file_section(body)

        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=10)

        # Options row
        self._options_section(body)

        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=10)

        # Output folder picker
        self._output_section(body)

        ttk.Separator(body, orient="horizontal").pack(fill="x", pady=10)

        # Generate button — uses _btn() helper for reliable macOS colours
        self._gen_btn = _btn(
            body, text="⚡  Generate Notices",
            command=self._on_generate,
            font=("Helvetica", 13, "bold"),
            bg=PRIMARY, fg=BTN_FG,
            padx=20, pady=10,
        )
        self._gen_btn.pack(pady=4)

        # Status log
        self._build_log(body)

        # Open folder button (hidden until notices generated)
        self._open_btn = _btn(
            body, text="📂  Open Output Folder",
            command=self._open_output_folder,
            bg=ACCENT, fg=BTN_FG,
            font=FONT_SUB, padx=12, pady=6,
        )
        # shown later

    def _file_section(self, parent):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=4)
        tk.Label(row, text="Excel Input File", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, width=20, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=self._excel_path, font=FONT_SUB,
                 fg=LABEL_FG, width=40, relief="solid", bd=1,
                 highlightthickness=1, highlightbackground=BORDER).pack(side="left", padx=6)
        _btn(row, "Browse…", self._browse_excel,
             bg=PRIMARY, fg=BTN_FG, font=FONT_SUB).pack(side="left")

    def _options_section(self, parent):
        frame = tk.Frame(parent, bg=BG)
        frame.pack(fill="x", pady=2)

        # Notice type
        tk.Label(frame, text="Notice Type", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, width=20, anchor="w").grid(row=0, column=0, pady=6, sticky="w")
        type_frame = tk.Frame(frame, bg=BG)
        type_frame.grid(row=0, column=1, sticky="w")
        for val, lbl in [("both",       "Both (Extension + Encashment)"),
                         ("extension",  "Extension Notice only"),
                         ("encashment", "Encashment Notice only")]:
            tk.Radiobutton(
                type_frame, text=lbl, variable=self._letter_type, value=val,
                font=FONT_SUB,
                bg=BG, fg=LABEL_FG,                  # ← explicit dark text
                activebackground=BG, activeforeground=LABEL_FG,
                selectcolor=CARD,                     # circle fill colour
            ).pack(side="left", padx=10)

        # Letter number
        tk.Label(frame, text="Letter Number", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, width=20, anchor="w").grid(row=1, column=0, pady=6, sticky="w")
        tk.Entry(frame, textvariable=self._letter_num, font=FONT_SUB,
                 fg=LABEL_FG, width=24, relief="solid", bd=1).grid(row=1, column=1, sticky="w", padx=4)

        # Letter date
        tk.Label(frame, text="Letter Date (YYYY-MM-DD)", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, width=20, anchor="w").grid(row=2, column=0, pady=6, sticky="w")
        tk.Entry(frame, textvariable=self._letter_date, font=FONT_SUB,
                 fg=LABEL_FG, width=16, relief="solid", bd=1).grid(row=2, column=1, sticky="w", padx=4)

        # Specific row (optional)
        tk.Label(frame, text="Specific Row (optional)", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, width=20, anchor="w").grid(row=3, column=0, pady=6, sticky="w")
        row_frame = tk.Frame(frame, bg=BG)
        row_frame.grid(row=3, column=1, sticky="w")
        tk.Entry(row_frame, textvariable=self._specific_row, font=FONT_SUB,
                 fg=LABEL_FG, width=8, relief="solid", bd=1).pack(side="left", padx=4)
        tk.Label(row_frame, text="Leave blank to generate for ALL due rows",
                 font=("Helvetica", 9), bg=BG, fg=MUTED_FG).pack(side="left")

    def _output_section(self, parent):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=4)
        tk.Label(row, text="Output Folder", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, width=20, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=self._output_path, font=FONT_SUB,
                 fg=LABEL_FG, width=40, relief="solid", bd=1).pack(side="left", padx=6)
        _btn(row, "Browse…", self._browse_output,
             bg=PRIMARY, fg=BTN_FG, font=FONT_SUB).pack(side="left")

    def _build_log(self, parent):
        tk.Label(parent, text="Log", font=FONT_BOLD,
                 bg=BG, fg=LABEL_FG, anchor="w").pack(fill="x", pady=(6, 2))
        log_frame = tk.Frame(parent, bg=BG)
        log_frame.pack(fill="both", expand=True)
        self._log = tk.Text(log_frame, height=9, font=FONT_MONO,
                            state="disabled", relief="solid", bd=1,
                            bg="#FAFAFA", wrap="word")
        sb = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=sb.set)
        self._log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Tag colours
        self._log.tag_configure("ok",      foreground=ACCENT)
        self._log.tag_configure("warn",    foreground="#E65100")
        self._log.tag_configure("err",     foreground=DANGER)
        self._log.tag_configure("info",    foreground=PRIMARY)
        self._log.tag_configure("section", foreground=PRIMARY, font=("Courier", 10, "bold"))

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _center_window(self, w, h):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select BGB Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self._excel_path.set(path)

    def _browse_output(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self._output_path.set(path)

    def _open_output_folder(self):
        folder = self._output_path.get()
        if os.path.isdir(folder):
            if sys.platform == "darwin":
                os.system(f'open "{folder}"')
            elif sys.platform == "win32":
                os.startfile(folder)
            else:
                os.system(f'xdg-open "{folder}"')

    def _log_write(self, text, tag=""):
        self._log.configure(state="normal")
        self._log.insert("end", text + "\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _log_clear(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", "end")
        self._log.configure(state="disabled")

    # ------------------------------------------------------------------
    # Generate — runs in a background thread so UI stays responsive
    # ------------------------------------------------------------------
    def _on_generate(self):
        excel = self._excel_path.get().strip()
        if not excel:
            messagebox.showerror("Missing Input", "Please select an Excel file first.")
            return
        if not os.path.isfile(excel):
            messagebox.showerror("File Not Found", f"Cannot find:\n{excel}")
            return

        # Validate date
        try:
            letter_date = date.fromisoformat(self._letter_date.get().strip())
        except ValueError:
            messagebox.showerror("Invalid Date",
                                 "Letter date must be in YYYY-MM-DD format.\n"
                                 "Example: 2026-05-05")
            return

        self._gen_btn.configure(state="disabled", text="⏳  Generating…",
                                bg="#6B8CBE", fg=BTN_FG)
        self._open_btn.pack_forget()
        self._log_clear()

        # Run in thread
        t = threading.Thread(target=self._run_generation,
                             args=(excel, letter_date), daemon=True)
        t.start()

    def _run_generation(self, excel_path: str, letter_date: date):
        today         = date.today()
        letter_type   = self._letter_type.get()
        letter_number = self._letter_num.get().strip()
        output_dir    = self._output_path.get().strip()
        specific_row  = self._specific_row.get().strip()

        try:
            specific_row_int = int(specific_row) if specific_row else None
        except ValueError:
            self.after(0, self._log_write, "⚠ 'Specific Row' must be a number or blank.", "warn")
            self.after(0, self._reset_button)
            return

        os.makedirs(output_dir, exist_ok=True)

        # ── Parse Excel ───────────────────────────────────────────────
        self.after(0, self._log_write, f"📄 Reading: {os.path.basename(excel_path)}", "info")
        try:
            rows = parse_excel(excel_path)
        except Exception as e:
            self.after(0, self._log_write, f"✗ Failed to read Excel: {e}", "err")
            self.after(0, self._reset_button)
            return

        self.after(0, self._log_write, f"   Found {len(rows)} data row(s).", "info")

        if specific_row_int:
            rows = [r for r in rows if r['row_index'] == specific_row_int]
            if not rows:
                self.after(0, self._log_write,
                           f"✗ Row {specific_row_int} not found.", "err")
                self.after(0, self._reset_button)
                return

        # ── Process rows ──────────────────────────────────────────────
        generated = skipped_exp = skipped_not = skipped_err = 0

        for row in rows:
            rid = row['row_index']

            missing = []
            if not row['ca_number']:     missing.append("CA number")
            if not row['bgb_number']:    missing.append("BGB number")
            if row['validity_date'] is None: missing.append("validity date")

            if missing:
                self.after(0, self._log_write,
                           f"⚠ Row {rid}: Missing {', '.join(missing)} — skipped.", "warn")
                skipped_err += 1
                continue

            validity = row['validity_date']

            if validity < today:
                self.after(0, self._log_write,
                           f"⚠ Row {rid}: BGB expired on {format_date(validity)} — skipped.", "warn")
                skipped_exp += 1
                continue

            if not validity_in_window(validity, today):
                skipped_not += 1
                continue

            self.after(0, self._log_write,
                       f"✔ Row {rid}: {row['contractor_name']} | "
                       f"valid upto {format_date(validity)}", "ok")
            try:
                if letter_type in ("extension", "both"):
                    fp = generate_extension_notice(
                        row=row, letter_number=letter_number,
                        letter_date=letter_date, output_dir=output_dir)
                    self.after(0, self._log_write,
                               f"   → Extension: {os.path.basename(fp)}", "ok")
                    generated += 1

                if letter_type in ("encashment", "both"):
                    fp = generate_encashment_notice(
                        row=row, letter_number=letter_number,
                        letter_date=letter_date,
                        prev_letter_number=letter_number,
                        prev_letter_date=letter_date,
                        output_dir=output_dir)
                    self.after(0, self._log_write,
                               f"   → Encashment: {os.path.basename(fp)}", "ok")
                    generated += 1

            except Exception as e:
                self.after(0, self._log_write,
                           f"   ✗ Error generating for row {rid}: {e}", "err")
                skipped_err += 1

        # ── Summary ───────────────────────────────────────────────────
        self.after(0, self._log_write, "─" * 55, "section")
        self.after(0, self._log_write,
                   f"DONE  |  Generated: {generated}  |  "
                   f"Expired: {skipped_exp}  |  Not due: {skipped_not}  |  "
                   f"Errors: {skipped_err}", "section")
        self.after(0, self._log_write, f"Output folder: {output_dir}", "info")

        if generated > 0:
            self.after(0, self._open_btn.pack, {"pady": 6})

        self.after(0, self._reset_button)

    def _reset_button(self):
        self._gen_btn.configure(
            state="normal", text="⚡  Generate Notices",
            bg=PRIMARY, fg=BTN_FG,
        )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    app = BGBApp()
    app.mainloop()
