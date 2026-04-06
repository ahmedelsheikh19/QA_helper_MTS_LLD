import customtkinter as ctk
from tkinter import filedialog, messagebox
import logging
import threading
import os
import sys
import webbrowser

# ── App-wide appearance ────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

FONT_TITLE  = ("Segoe UI", 22, "bold")
FONT_LABEL  = ("Segoe UI", 13)
FONT_SMALL  = ("Segoe UI", 11)
FONT_LOG    = ("Consolas", 11)
FONT_BTN    = ("Segoe UI", 13, "bold")

COLOR_SUCCESS = "#2ecc71"
COLOR_DANGER  = "#e74c3c"
COLOR_ACCENT  = "#3498db"
COLOR_CARD    = "#1e1e2e"
COLOR_BG      = "#13131f"
COLOR_SIDEBAR = "#16162a"


# ── Logging → CTkTextbox ───────────────────────────────────────────────────
class TextHandler(logging.Handler):
    COLORS = {
        "INFO":    "#c8d6e5",
        "WARNING": "#f9ca24",
        "ERROR":   "#e74c3c",
        "DEBUG":   "#7f8c8d",
    }

    def __init__(self, textbox: ctk.CTkTextbox):
        super().__init__()
        self.tb = textbox
        self.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s",
                                            datefmt="%H:%M:%S"))

    def emit(self, record):
        msg   = self.format(record) + "\n"
        color = self.COLORS.get(record.levelname, "#c8d6e5")

        def _insert():
            self.tb.configure(state="normal")
            self.tb.insert("end", msg)
            # naive colorise last inserted line via tag — CTkTextbox wraps tk.Text
            self.tb._textbox.tag_add(record.levelname,
                                     f"end - {len(msg)+1}c", "end - 1c")
            self.tb._textbox.tag_config(record.levelname, foreground=color)
            self.tb.see("end")
            self.tb.configure(state="disabled")

        self.tb.after(0, _insert)


# ── Reusable file-picker row ───────────────────────────────────────────────
class FilePickerRow(ctk.CTkFrame):
    def __init__(self, parent, label, multi=False, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.multi = multi
        self.columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text=label, font=FONT_SMALL,
                     text_color="#8899aa").grid(row=0, column=0, columnspan=2,
                                               sticky="w", pady=(0, 4))

        if not multi:
            self._var = ctk.StringVar()
            self._entry = ctk.CTkEntry(self, textvariable=self._var,
                                       font=FONT_SMALL, height=36,
                                       placeholder_text="Select file…")
            self._entry.grid(row=1, column=0, sticky="ew", padx=(0, 8))
            ctk.CTkButton(self, text="Browse", width=90, height=36,
                          font=FONT_SMALL, fg_color="#2d2d44",
                          hover_color="#3d3d5c",
                          command=self._browse_single).grid(row=1, column=1)
        else:
            self._listbox_var = []
            self._lb = ctk.CTkTextbox(self, height=90, font=FONT_SMALL,
                                      state="disabled", fg_color="#0d0d1a",
                                      text_color="#c8d6e5")
            self._lb.grid(row=1, column=0, sticky="ew", padx=(0, 8))
            btn_col = ctk.CTkFrame(self, fg_color="transparent")
            btn_col.grid(row=1, column=1)
            ctk.CTkButton(btn_col, text="Add", width=80, height=34,
                          font=FONT_SMALL, fg_color="#2d2d44",
                          hover_color="#3d3d5c",
                          command=self._add).pack(pady=(0, 6))
            ctk.CTkButton(btn_col, text="Clear", width=80, height=34,
                          font=FONT_SMALL, fg_color="#2d2d44",
                          hover_color="#3d3d5c",
                          command=self._clear).pack()

    def _browse_single(self):
        p = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if p:
            self._var.set(p)

    def _add(self):
        paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        for p in paths:
            if p not in self._listbox_var:
                self._listbox_var.append(p)
        self._refresh_lb()

    def _clear(self):
        self._listbox_var.clear()
        self._refresh_lb()

    def _refresh_lb(self):
        self._lb.configure(state="normal")
        self._lb.delete("1.0", "end")
        for p in self._listbox_var:
            self._lb.insert("end", f"  {os.path.basename(p)}\n")
        self._lb.configure(state="disabled")

    def get(self):
        if self.multi:
            return list(self._listbox_var)
        return self._var.get().strip()


# ── Output picker row ──────────────────────────────────────────────────────
class OutputRow(ctk.CTkFrame):
    def __init__(self, parent, default_name, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.columnconfigure(0, weight=1)
        ctk.CTkLabel(self, text="Output HTML file", font=FONT_SMALL,
                     text_color="#8899aa").grid(row=0, column=0, columnspan=2,
                                               sticky="w", pady=(0, 4))
        self._var = ctk.StringVar(value=default_name)
        ctk.CTkEntry(self, textvariable=self._var, font=FONT_SMALL,
                     height=36).grid(row=1, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkButton(self, text="Save as", width=90, height=36,
                      font=FONT_SMALL, fg_color="#2d2d44",
                      hover_color="#3d3d5c",
                      command=self._browse).grid(row=1, column=1)

    def _browse(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".html",
            initialfile=self._var.get(),
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")])
        if p:
            self._var.set(p)

    def get(self):
        return self._var.get().strip()


# ── Shared log + run section ───────────────────────────────────────────────
class RunSection(ctk.CTkFrame):
    def __init__(self, parent, btn_label, btn_color, on_run, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.columnconfigure(0, weight=1)

        self.run_btn = ctk.CTkButton(
            self, text=btn_label, font=FONT_BTN,
            height=44, corner_radius=10,
            fg_color=btn_color, hover_color=self._darken(btn_color),
            command=on_run)
        self.run_btn.grid(row=0, column=0, sticky="ew", pady=(0, 12))

        self.progress = ctk.CTkProgressBar(self, height=6, corner_radius=4)
        self.progress.set(0)
        self.progress.grid(row=1, column=0, sticky="ew", pady=(0, 12))

        ctk.CTkLabel(self, text="Log output", font=FONT_SMALL,
                     text_color="#8899aa").grid(row=2, column=0, sticky="w")
        self.log_box = ctk.CTkTextbox(self, height=180, font=FONT_LOG,
                                      state="disabled", fg_color="#0d0d1a",
                                      text_color="#c8d6e5", corner_radius=8)
        self.log_box.grid(row=3, column=0, sticky="nsew", pady=(4, 0))
        self.rowconfigure(3, weight=1)

    @staticmethod
    def _darken(hex_color):
        r, g, b = int(hex_color[1:3], 16), int(hex_color[3:5], 16), int(hex_color[5:7], 16)
        r, g, b = max(0, r - 30), max(0, g - 30), max(0, b - 30)
        return f"#{r:02x}{g:02x}{b:02x}"

    def set_running(self, running: bool):
        if running:
            self.run_btn.configure(state="disabled", text="Running…")
            self.progress.configure(mode="indeterminate")
            self.progress.start()
        else:
            self.run_btn.configure(state="normal",
                                   text=self.run_btn._text)
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(1)

    def clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def handler(self):
        return TextHandler(self.log_box)


# ── LLD panel ─────────────────────────────────────────────────────────────
class LLDPanel(ctk.CTkScrollableFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=COLOR_BG, **kw)
        self.columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="LLD Traceability", font=FONT_TITLE,
                     text_color="#e0e6f0").grid(row=0, column=0, sticky="w",
                                               pady=(0, 4))
        ctk.CTkLabel(self, text="Requirements  →  Low-Level Design",
                     font=FONT_SMALL, text_color="#6677aa").grid(
            row=1, column=0, sticky="w", pady=(0, 20))

        card = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12)
        card.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        card.columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text="Input Files", font=("Segoe UI", 13, "bold"),
                     text_color="#aabbcc").pack(anchor="w", padx=20, pady=(16, 8))

        self.req_row = FilePickerRow(card, "Requirements Excel file (.xlsx)")
        self.req_row.pack(fill="x", padx=20, pady=(0, 12))
        self.lld_row = FilePickerRow(card, "LLD Excel file (.xlsx)")
        self.lld_row.pack(fill="x", padx=20, pady=(0, 16))

        card2 = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12)
        card2.grid(row=3, column=0, sticky="ew", pady=(0, 16))
        card2.columnconfigure(0, weight=1)
        ctk.CTkLabel(card2, text="Output", font=("Segoe UI", 13, "bold"),
                     text_color="#aabbcc").pack(anchor="w", padx=20, pady=(16, 8))
        self.out_row = OutputRow(card2, "traceability_report.html")
        self.out_row.pack(fill="x", padx=20, pady=(0, 16))

        self.run_sec = RunSection(self, btn_label="⚡  Generate LLD Report",
                                  btn_color="#1a6eb5", on_run=self._run)
        self.run_sec.grid(row=4, column=0, sticky="nsew", pady=(0, 16))
        self.rowconfigure(4, weight=1)

    def _run(self):
        req_file = self.req_row.get()
        lld_file = self.lld_row.get()
        out_file = self.out_row.get()

        if not req_file or not lld_file:
            messagebox.showerror("Missing input",
                                 "Please select both the Requirements and LLD files.")
            return
        for label, path in [("Requirements", req_file), ("LLD", lld_file)]:
            if not os.path.exists(path):
                messagebox.showerror("File not found",
                                     f"{label} file not found:\n{path}")
                return

        self.run_sec.clear_log()
        self.run_sec.set_running(True)
        threading.Thread(target=self._worker,
                         args=(req_file, lld_file, out_file),
                         daemon=True).start()

    def _worker(self, req_file, lld_file, out_file):
        logger  = logging.getLogger()
        handler = self.run_sec.handler()
        logger.addHandler(handler)
        try:
            from generate_traceabilityLLD import (
                extract_reqs_from_req_file,
                extract_reqs_from_lld_file,
                generate_html_report,
            )
            req_ids  = extract_reqs_from_req_file(req_file)
            lld_reqs = extract_reqs_from_lld_file(lld_file)
            if not req_ids and not lld_reqs:
                logging.error("No requirements found in both files.")
            else:
                generate_html_report(req_ids, lld_reqs, out_file)
                abs_out = os.path.abspath(out_file)
                self.after(0, lambda: _prompt_open(abs_out))
        except Exception as e:
            logging.error(f"Unexpected error: {e}")
        finally:
            logger.removeHandler(handler)
            self.after(0, lambda: self.run_sec.set_running(False))


# ── UT panel ──────────────────────────────────────────────────────────────
class UTPanel(ctk.CTkScrollableFrame):
    def __init__(self, parent, **kw):
        super().__init__(parent, fg_color=COLOR_BG, **kw)
        self.columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="UT Traceability", font=FONT_TITLE,
                     text_color="#e0e6f0").grid(row=0, column=0, sticky="w",
                                               pady=(0, 4))
        ctk.CTkLabel(self, text="Requirements  →  Unit Tests",
                     font=FONT_SMALL, text_color="#6677aa").grid(
            row=1, column=0, sticky="w", pady=(0, 20))

        card = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12)
        card.grid(row=2, column=0, sticky="ew", pady=(0, 16))
        card.columnconfigure(0, weight=1)
        ctk.CTkLabel(card, text="Input Files", font=("Segoe UI", 13, "bold"),
                     text_color="#aabbcc").pack(anchor="w", padx=20, pady=(16, 8))

        self.req_row = FilePickerRow(card, "Requirements Excel file (.xlsx)")
        self.req_row.pack(fill="x", padx=20, pady=(0, 12))
        self.ut_row  = FilePickerRow(card, "UT Excel file(s) — add one or more",
                                     multi=True)
        self.ut_row.pack(fill="x", padx=20, pady=(0, 16))

        card2 = ctk.CTkFrame(self, fg_color=COLOR_CARD, corner_radius=12)
        card2.grid(row=3, column=0, sticky="ew", pady=(0, 16))
        card2.columnconfigure(0, weight=1)
        ctk.CTkLabel(card2, text="Output", font=("Segoe UI", 13, "bold"),
                     text_color="#aabbcc").pack(anchor="w", padx=20, pady=(16, 8))
        self.out_row = OutputRow(card2, "ut_traceability_report.html")
        self.out_row.pack(fill="x", padx=20, pady=(0, 16))

        self.run_sec = RunSection(self, btn_label="⚡  Generate UT Report",
                                  btn_color="#176b3a", on_run=self._run)
        self.run_sec.grid(row=4, column=0, sticky="nsew", pady=(0, 16))
        self.rowconfigure(4, weight=1)

    def _run(self):
        req_file = self.req_row.get()
        ut_files = self.ut_row.get()
        out_file = self.out_row.get()

        if not req_file or not ut_files:
            messagebox.showerror("Missing input",
                                 "Please select the Requirements file and at least one UT file.")
            return
        if not os.path.exists(req_file):
            messagebox.showerror("File not found",
                                 f"Requirements file not found:\n{req_file}")
            return
        for f in ut_files:
            if not os.path.exists(f):
                messagebox.showerror("File not found", f"UT file not found:\n{f}")
                return

        self.run_sec.clear_log()
        self.run_sec.set_running(True)
        threading.Thread(target=self._worker,
                         args=(req_file, ut_files, out_file),
                         daemon=True).start()

    def _worker(self, req_file, ut_files, out_file):
        logger  = logging.getLogger()
        handler = self.run_sec.handler()
        logger.addHandler(handler)
        try:
            from generate_ut_traceability import (
                extract_reqs_from_req_file,
                extract_reqs_from_ut_files,
                generate_html_report,
            )
            req_ids             = extract_reqs_from_req_file(req_file)
            ut_reqs, fmt_errors = extract_reqs_from_ut_files(ut_files)
            if not req_ids and not ut_reqs:
                logging.error("No requirements found in both files.")
            else:
                generate_html_report(req_ids, ut_reqs, fmt_errors, out_file)
                abs_out = os.path.abspath(out_file)
                self.after(0, lambda: _prompt_open(abs_out))
        except Exception as e:
            logging.error(f"Unexpected error: {e}")
        finally:
            logger.removeHandler(handler)
            self.after(0, lambda: self.run_sec.set_running(False))


# ── Sidebar nav ───────────────────────────────────────────────────────────
class Sidebar(ctk.CTkFrame):
    def __init__(self, parent, on_select, **kw):
        super().__init__(parent, width=200, corner_radius=0,
                         fg_color=COLOR_SIDEBAR, **kw)
        self.pack_propagate(False)
        self._btns = {}
        self._on_select = on_select

        ctk.CTkLabel(self, text="Traceability\nGenerator",
                     font=("Segoe UI", 15, "bold"),
                     text_color="#c8d6e5",
                     justify="center").pack(pady=(28, 32), padx=16)

        for key, label, icon in [
            ("lld", "LLD Report",  "📐"),
            ("ut",  "UT Report",   "🧪"),
        ]:
            btn = ctk.CTkButton(
                self, text=f"  {icon}  {label}", anchor="w",
                font=("Segoe UI", 13), height=44,
                fg_color="transparent", hover_color="#2a2a45",
                text_color="#aabbdd", corner_radius=8,
                command=lambda k=key: self._select(k))
            btn.pack(fill="x", padx=10, pady=4)
            self._btns[key] = btn

        # appearance toggle at bottom
        ctk.CTkLabel(self, text="").pack(expand=True)
        self._mode_btn = ctk.CTkButton(
            self, text="☀  Light mode", anchor="w",
            font=FONT_SMALL, height=36, fg_color="transparent",
            hover_color="#2a2a45", text_color="#8899aa",
            command=self._toggle_mode)
        self._mode_btn.pack(fill="x", padx=10, pady=(0, 16))

        self._current = None
        self._select("lld")

    def _select(self, key):
        if self._current:
            self._btns[self._current].configure(fg_color="transparent",
                                                text_color="#aabbdd")
        self._current = key
        self._btns[key].configure(fg_color="#2a2a45", text_color="#ffffff")
        self._on_select(key)

    def _toggle_mode(self):
        mode = ctk.get_appearance_mode()
        if mode == "Dark":
            ctk.set_appearance_mode("light")
            self._mode_btn.configure(text="🌙  Dark mode")
        else:
            ctk.set_appearance_mode("dark")
            self._mode_btn.configure(text="☀  Light mode")


# ── Helpers ────────────────────────────────────────────────────────────────
def _prompt_open(path):
    if messagebox.askyesno("Done ✓", f"Report saved to:\n{path}\n\nOpen in browser?"):
        webbrowser.open(f"file:///{path}")


# ── Main window ───────────────────────────────────────────────────────────
def main():
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    logging.getLogger().setLevel(logging.INFO)

    root = ctk.CTk()
    root.title("Traceability Generator")
    root.geometry("960x680")
    root.minsize(800, 560)

    root.grid_rowconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)

    panels = {}

    def show(key):
        for k, p in panels.items():
            if k == key:
                p.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
            else:
                p.grid_remove()

    lld_panel = LLDPanel(root)
    ut_panel  = UTPanel(root)
    panels["lld"] = lld_panel
    panels["ut"]  = ut_panel

    sidebar = Sidebar(root, on_select=show)
    sidebar.grid(row=0, column=0, sticky="ns")

    root.mainloop()


if __name__ == "__main__":
    main()
