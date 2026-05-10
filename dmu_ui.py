"""Simple preview UI for DMU result files."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk


REPO_ROOT = Path(__file__).resolve().parents[1]
RESULT_ROOT = REPO_ROOT / "result" / "dmu_agent"
DMU_AGENT_PATH = REPO_ROOT / "catia_agents" / "DMU Agent.py"
MODE_BETWEEN_ALL = "between_all_components"
MODE_SELECTION_AGAINST_ALL = "selection_against_all"
MODE_BETWEEN_TWO = "between_two_components"


def parse_args():
    parser = argparse.ArgumentParser(description="Run CATIA DMU analysis and open the preview UI.")
    parser.add_argument("--viewer-only", action="store_true", help="Skip PowerShell prompts and open the preview UI only.")
    parser.add_argument("--mode", choices=[MODE_BETWEEN_ALL, MODE_SELECTION_AGAINST_ALL, MODE_BETWEEN_TWO], default=None)
    parser.add_argument("--clearance_mm", type=float, default=None)
    parser.add_argument("--part_number", type=str, default=None)
    parser.add_argument("--part_number_a", type=str, default=None)
    parser.add_argument("--part_number_b", type=str, default=None)
    return parser.parse_args()


def choose_mode_interactive():
    print("")
    print("DMU analysis mode:")
    print("  1. Between all components")
    print("  2. Selection against all")
    print("  3. Between two components")
    while True:
        choice = input("Choose mode [1-3, default 1]: ").strip()
        if choice in ("", "1"):
            return MODE_BETWEEN_ALL
        if choice == "2":
            return MODE_SELECTION_AGAINST_ALL
        if choice == "3":
            return MODE_BETWEEN_TWO
        print("Please enter 1, 2, or 3.")


def prompt_required_text(prompt_text: str) -> str:
    while True:
        value = input(prompt_text).strip()
        if value:
            return value
        print("Please enter a value.")


def prompt_clearance(default_value: float = 5.0) -> float:
    while True:
        raw = input("Enter clearance in mm [default {}]: ".format(default_value)).strip()
        if not raw:
            return float(default_value)
        try:
            return float(raw)
        except Exception:
            print("Please enter a valid numeric clearance.")


def prompt_startup_run_options():
    print("")
    print("DMU startup:")
    print("  1. Run DMU analysis, then open preview UI")
    print("  2. Open latest preview UI only")
    while True:
        choice = input("Choose startup mode [1-2, default 1]: ").strip()
        if choice in ("", "1"):
            mode = choose_mode_interactive()
            clearance_mm = prompt_clearance(5.0)
            part_number = ""
            part_number_a = ""
            part_number_b = ""
            if mode == MODE_SELECTION_AGAINST_ALL:
                part_number = prompt_required_text("Enter part number for selection against all: ")
            elif mode == MODE_BETWEEN_TWO:
                part_number_a = prompt_required_text("Enter first part number: ")
                part_number_b = prompt_required_text("Enter second part number: ")
            return {
                "run_analysis": True,
                "mode": mode,
                "clearance_mm": clearance_mm,
                "part_number": part_number,
                "part_number_a": part_number_a,
                "part_number_b": part_number_b,
            }
        if choice == "2":
            return {"run_analysis": False}
        print("Please enter 1 or 2.")


def build_agent_command(mode: str, clearance_mm: str | float, part_number: str, part_number_a: str, part_number_b: str):
    command = [
        sys.executable,
        str(DMU_AGENT_PATH),
        "--mode",
        mode,
        "--clearance_mm",
        str(clearance_mm),
        "--show_rows",
        "200",
    ]
    if mode == MODE_SELECTION_AGAINST_ALL and part_number:
        command.extend(["--part_number", part_number])
    if mode == MODE_BETWEEN_TWO:
        command.extend(["--part_number_a", part_number_a, "--part_number_b", part_number_b])
    return command


def find_latest_json() -> Path | None:
    if not RESULT_ROOT.exists():
        return None
    candidates = sorted(RESULT_ROOT.glob("*/results.json"), key=lambda path: path.stat().st_mtime, reverse=True)
    return candidates[0] if candidates else None


class DMUPreviewApp:
    def __init__(self, root: tk.Tk, initial_json_path: Path | None = None):
        self.root = root
        self.root.title("DMU Result Preview")
        self.root.geometry("1500x920")

        self.status_var = tk.StringVar(value="Load a DMU results.json file or use Load Latest.")
        self.path_var = tk.StringVar(value="")
        self.filter_var = tk.StringVar(value="all")
        self.row_var = tk.StringVar(value="Row: - / -")
        self.run_in_progress = False

        self.current_json_path: Path | None = None
        self.all_results: list[dict] = []
        self.filtered_results: list[dict] = []
        self.preview_image = None
        self.preview_pil_image: Image.Image | None = None
        self.preview_zoom = 1.0
        self.preview_pan_x = 0.0
        self.preview_pan_y = 0.0
        self.popup = None
        self.popup_label = None
        self.popup_image = None
        self.popup_zoom = 1.0
        self.popup_pan_x = 0.0
        self.popup_pan_y = 0.0
        self._drag_target = None
        self._drag_start_x = 0
        self._drag_start_y = 0

        self._build_ui()
        if initial_json_path is not None and initial_json_path.exists():
            self._load_json(initial_json_path)
        else:
            self._load_latest(silent=True)

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        toolbar = ttk.Frame(main)
        toolbar.pack(fill="x")

        ttk.Button(toolbar, text="Run DMU", command=self._open_run_dialog).pack(side="left")
        ttk.Button(toolbar, text="Load Latest", command=self._load_latest).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Open Results JSON", command=self._open_results_file).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Open Run Folder", command=self._open_run_folder).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Open Image", command=self._open_selected_image).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Large Preview", command=self._open_large_preview).pack(side="left", padx=(8, 0))
        ttk.Button(toolbar, text="Zoom In", command=lambda: self._zoom_preview(1.2)).pack(side="left", padx=(12, 0))
        ttk.Button(toolbar, text="Zoom Out", command=lambda: self._zoom_preview(1 / 1.2)).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Reset Zoom", command=self._reset_zoom).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="Previous", command=lambda: self._move_selection(-1)).pack(side="left", padx=(20, 0))
        ttk.Button(toolbar, text="Next", command=lambda: self._move_selection(1)).pack(side="left", padx=(6, 0))
        ttk.Label(toolbar, textvariable=self.row_var).pack(side="left", padx=(12, 0))
        ttk.Label(toolbar, text="Filter").pack(side="right")

        self.filter_combo = ttk.Combobox(
            toolbar,
            textvariable=self.filter_var,
            state="readonly",
            width=16,
            values=["all", "contact", "clash", "clearance"],
        )
        self.filter_combo.pack(side="right", padx=(0, 8))
        self.filter_combo.bind("<<ComboboxSelected>>", lambda _event: self._apply_filter())

        ttk.Label(main, textvariable=self.status_var).pack(anchor="w", pady=(10, 0))
        ttk.Label(main, textvariable=self.path_var, foreground="#555555").pack(anchor="w", pady=(2, 8))

        paned = ttk.Panedwindow(main, orient="horizontal")
        paned.pack(fill="both", expand=True)

        left = ttk.Frame(paned)
        right = ttk.Frame(paned)
        paned.add(left, weight=3)
        paned.add(right, weight=2)

        columns = ("No", "Product1", "Product2", "Type", "Value", "Status", "Info")
        self.tree = ttk.Treeview(left, columns=columns, show="headings", height=30)
        self.tree.pack(side="left", fill="both", expand=True)

        scroll_y = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        scroll_y.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll_y.set)

        widths = {
            "No": 55,
            "Product1": 220,
            "Product2": 240,
            "Type": 95,
            "Value": 85,
            "Status": 120,
            "Info": 90,
        }
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=widths[column], anchor="w", stretch=True)

        self.tree.bind("<<TreeviewSelect>>", lambda _event: self._show_preview())
        self.tree.bind("<Double-1>", lambda _event: self._open_large_preview())

        preview_frame = ttk.LabelFrame(right, text="Preview", padding=10)
        preview_frame.pack(fill="both", expand=True)
        preview_frame.bind("<Configure>", lambda _event: self._render_main_preview())

        self.preview_label = ttk.Label(preview_frame, text="Select a result row to see the preview.", anchor="center")
        self.preview_label.pack(fill="both", expand=True)
        self.preview_label.bind("<MouseWheel>", self._on_mousewheel)
        self.preview_label.bind("<Button-4>", self._on_mousewheel)
        self.preview_label.bind("<Button-5>", self._on_mousewheel)
        self.preview_label.bind("<ButtonPress-1>", lambda event: self._start_pan("main", event))
        self.preview_label.bind("<B1-Motion>", lambda event: self._drag_pan("main", event))
        self.preview_label.bind("<ButtonRelease-1>", lambda _event: self._end_pan())

        self.preview_meta = tk.Text(right, height=10, wrap="word")
        self.preview_meta.pack(fill="x", pady=(8, 0))
        self.preview_meta.configure(state="disabled")

    def _find_latest_json(self) -> Path | None:
        return find_latest_json()

    def _load_latest(self, silent: bool = False):
        json_path = self._find_latest_json()
        if json_path is None:
            if not silent:
                messagebox.showinfo("No Results", "No DMU results were found under result/dmu_agent.")
            return
        self._load_json(json_path)

    def _open_results_file(self):
        path = filedialog.askopenfilename(
            title="Open DMU results.json",
            initialdir=str(RESULT_ROOT if RESULT_ROOT.exists() else REPO_ROOT),
            filetypes=[("JSON files", "*.json")],
        )
        if not path:
            return
        self._load_json(Path(path))

    def _load_json(self, json_path: Path):
        payload = json.loads(json_path.read_text(encoding="utf-8"))
        self.current_json_path = json_path
        self.all_results = list(payload.get("results", []))
        self.path_var.set(str(json_path))
        self.status_var.set("Loaded {} result rows.".format(len(self.all_results)))
        self._apply_filter()

    def _open_run_dialog(self):
        if self.run_in_progress:
            messagebox.showinfo("DMU Running", "A DMU run is already in progress.")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Run DMU Analysis")
        dialog.geometry("520x270")
        dialog.transient(self.root)
        dialog.grab_set()

        mode_var = tk.StringVar(value=MODE_BETWEEN_ALL)
        clearance_var = tk.StringVar(value="5")
        part_number_var = tk.StringVar(value="")
        part_number_a_var = tk.StringVar(value="")
        part_number_b_var = tk.StringVar(value="")

        frame = ttk.Frame(dialog, padding=14)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Mode").grid(row=0, column=0, sticky="w", pady=(0, 8))
        mode_combo = ttk.Combobox(
            frame,
            textvariable=mode_var,
            state="readonly",
            width=32,
            values=[MODE_BETWEEN_ALL, MODE_SELECTION_AGAINST_ALL, MODE_BETWEEN_TWO],
        )
        mode_combo.grid(row=0, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Clearance (mm)").grid(row=1, column=0, sticky="w", pady=(0, 8))
        clearance_entry = ttk.Entry(frame, textvariable=clearance_var, width=18)
        clearance_entry.grid(row=1, column=1, sticky="w", pady=(0, 8))

        ttk.Label(frame, text="Part Number").grid(row=2, column=0, sticky="w", pady=(0, 8))
        part_number_entry = ttk.Entry(frame, textvariable=part_number_var, width=36)
        part_number_entry.grid(row=2, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Part Number A").grid(row=3, column=0, sticky="w", pady=(0, 8))
        part_number_a_entry = ttk.Entry(frame, textvariable=part_number_a_var, width=36)
        part_number_a_entry.grid(row=3, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(frame, text="Part Number B").grid(row=4, column=0, sticky="w", pady=(0, 8))
        part_number_b_entry = ttk.Entry(frame, textvariable=part_number_b_var, width=36)
        part_number_b_entry.grid(row=4, column=1, sticky="ew", pady=(0, 8))

        help_var = tk.StringVar()
        ttk.Label(frame, textvariable=help_var, foreground="#555555", wraplength=460).grid(
            row=5, column=0, columnspan=2, sticky="w", pady=(8, 12)
        )

        button_row = ttk.Frame(frame)
        button_row.grid(row=6, column=0, columnspan=2, sticky="e")

        frame.columnconfigure(1, weight=1)

        def refresh_mode_fields(*_args):
            mode = mode_var.get().strip()
            is_selection = mode == MODE_SELECTION_AGAINST_ALL
            is_between_two = mode == MODE_BETWEEN_TWO

            part_number_entry.configure(state="normal" if is_selection else "disabled")
            part_number_a_entry.configure(state="normal" if is_between_two else "disabled")
            part_number_b_entry.configure(state="normal" if is_between_two else "disabled")

            if mode == MODE_BETWEEN_ALL:
                help_var.set("Runs directly on the active CATProduct. No part number is needed.")
            elif mode == MODE_SELECTION_AGAINST_ALL:
                help_var.set("Enter one part number. The agent will compare that part against all others.")
            else:
                help_var.set("Enter two part numbers. The agent will compare those two selected components.")

        mode_var.trace_add("write", refresh_mode_fields)
        refresh_mode_fields()

        def submit():
            mode = mode_var.get().strip()
            clearance_text = clearance_var.get().strip()

            try:
                float(clearance_text)
            except Exception:
                messagebox.showerror("Invalid Clearance", "Enter a valid numeric clearance value.")
                return

            if mode == MODE_SELECTION_AGAINST_ALL and not part_number_var.get().strip():
                messagebox.showerror("Missing Part Number", "Enter a part number for selection-against-all.")
                return
            if mode == MODE_BETWEEN_TWO:
                if not part_number_a_var.get().strip() or not part_number_b_var.get().strip():
                    messagebox.showerror("Missing Part Numbers", "Enter both part numbers for between-two-components.")
                    return

            dialog.destroy()
            self._run_analysis(
                mode=mode,
                clearance_mm=clearance_text,
                part_number=part_number_var.get().strip(),
                part_number_a=part_number_a_var.get().strip(),
                part_number_b=part_number_b_var.get().strip(),
            )

        ttk.Button(button_row, text="Cancel", command=dialog.destroy).pack(side="right")
        ttk.Button(button_row, text="Run", command=submit).pack(side="right", padx=(0, 8))

    def _run_analysis(self, mode: str, clearance_mm: str, part_number: str, part_number_a: str, part_number_b: str):
        if not DMU_AGENT_PATH.exists():
            messagebox.showerror("DMU Agent Missing", "Could not find {}".format(DMU_AGENT_PATH))
            return

        command = build_agent_command(mode, clearance_mm, part_number, part_number_a, part_number_b)

        self.run_in_progress = True
        self.status_var.set("Running DMU analysis...")
        self.path_var.set("")

        def worker():
            completed = subprocess.run(
                command,
                cwd=str(REPO_ROOT),
                capture_output=True,
                text=True,
            )
            self.root.after(0, lambda: self._finish_analysis_run(completed))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_analysis_run(self, completed: subprocess.CompletedProcess[str]):
        self.run_in_progress = False
        if completed.returncode != 0:
            stderr_text = completed.stderr.strip() or completed.stdout.strip() or "DMU analysis failed."
            self.status_var.set("DMU analysis failed.")
            messagebox.showerror("DMU Run Failed", stderr_text)
            return

        latest = self._find_latest_json()
        if latest is None:
            self.status_var.set("DMU run finished, but no results.json was found.")
            messagebox.showwarning("No Results", "DMU run finished, but no results.json was found.")
            return

        self._load_json(latest)
        self.status_var.set("DMU analysis finished and latest results were loaded.")

    def _apply_filter(self):
        active_filter = self.filter_var.get().strip().lower()
        if active_filter == "all":
            self.filtered_results = list(self.all_results)
        else:
            self.filtered_results = [row for row in self.all_results if str(row.get("Type", "")).strip().lower() == active_filter]

        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in self.filtered_results:
            values = (
                row.get("No", ""),
                row.get("Product1", ""),
                row.get("Product2", ""),
                row.get("Type", ""),
                row.get("Value", ""),
                row.get("Status", ""),
                row.get("Info", ""),
            )
            self.tree.insert("", "end", iid=str(row.get("No")), values=values)

        if self.filtered_results:
            first = str(self.filtered_results[0].get("No"))
            self.tree.selection_set(first)
            self.tree.focus(first)
            self.tree.see(first)
            self._show_preview()
        else:
            self.row_var.set("Row: 0 / 0")
            self._clear_preview()

    def _selected_row(self) -> dict | None:
        selected = self.tree.selection()
        if not selected:
            return None
        selected_no = selected[0]
        return next((row for row in self.filtered_results if str(row.get("No")) == selected_no), None)

    def _show_preview(self):
        row = self._selected_row()
        if row is None:
            self._clear_preview()
            return

        index = next((i for i, item in enumerate(self.filtered_results, start=1) if item.get("No") == row.get("No")), 0)
        self.row_var.set("Row: {} / {}".format(index, len(self.filtered_results)))

        image_path = Path(str(row.get("Image", "")))
        if not image_path.exists():
            self.preview_label.configure(text="Preview image not found:\n{}".format(image_path), image="")
            self.preview_image = None
            return

        with Image.open(image_path) as image:
            self.preview_pil_image = image.copy()
        self.preview_zoom = 1.0
        self.preview_pan_x = 0.0
        self.preview_pan_y = 0.0

        self._render_main_preview()

        meta = [
            "No: {}".format(row.get("No", "")),
            "Product1: {}".format(row.get("Product1", "")),
            "Product2: {}".format(row.get("Product2", "")),
            "Type: {}".format(row.get("Type", "")),
            "Value: {}".format(row.get("Value", "")),
            "Status: {}".format(row.get("Status", "")),
            "Info: {}".format(row.get("Info", "")),
            "Image: {}".format(image_path),
        ]
        self.preview_meta.configure(state="normal")
        self.preview_meta.delete("1.0", "end")
        self.preview_meta.insert("1.0", "\n".join(meta))
        self.preview_meta.configure(state="disabled")

        if self.popup is not None and self.popup.winfo_exists():
            self._refresh_popup()

    def _clear_preview(self):
        self.preview_label.configure(text="Select a result row to see the preview.", image="")
        self.preview_image = None
        self.preview_pil_image = None
        self.preview_zoom = 1.0
        self.preview_pan_x = 0.0
        self.preview_pan_y = 0.0
        self.preview_meta.configure(state="normal")
        self.preview_meta.delete("1.0", "end")
        self.preview_meta.configure(state="disabled")

    def _move_selection(self, delta: int):
        children = self.tree.get_children()
        if not children:
            return
        selected = self.tree.selection()
        if not selected:
            target_index = 0
        else:
            try:
                current_index = children.index(selected[0])
            except ValueError:
                current_index = 0
            target_index = max(0, min(len(children) - 1, current_index + delta))
        target = children[target_index]
        self.tree.selection_set(target)
        self.tree.focus(target)
        self.tree.see(target)
        self._show_preview()

    def _open_run_folder(self):
        if self.current_json_path is None:
            return
        os.startfile(str(self.current_json_path.parent))

    def _open_selected_image(self):
        row = self._selected_row()
        if row is None:
            return
        image_path = Path(str(row.get("Image", "")))
        if image_path.exists():
            os.startfile(str(image_path))

    def _open_large_preview(self):
        if self.preview_pil_image is None:
            return
        if self.popup is None or not self.popup.winfo_exists():
            self.popup = tk.Toplevel(self.root)
            self.popup.title("Preview")
            self.popup.geometry("1200x900")
            container = ttk.Frame(self.popup, padding=10)
            container.pack(fill="both", expand=True)
            self.popup_label = ttk.Label(container, anchor="center")
            self.popup_label.pack(fill="both", expand=True)
            self.popup.bind("<Configure>", lambda _event: self._refresh_popup())
            self.popup.bind("<MouseWheel>", self._on_popup_mousewheel)
            self.popup.bind("<Button-4>", self._on_popup_mousewheel)
            self.popup.bind("<Button-5>", self._on_popup_mousewheel)
            self.popup_label.bind("<ButtonPress-1>", lambda event: self._start_pan("popup", event))
            self.popup_label.bind("<B1-Motion>", lambda event: self._drag_pan("popup", event))
            self.popup_label.bind("<ButtonRelease-1>", lambda _event: self._end_pan())
        self.popup_zoom = max(self.popup_zoom, self.preview_zoom)
        self._refresh_popup()
        self.popup.deiconify()
        self.popup.lift()

    def _refresh_popup(self):
        if self.popup is None or not self.popup.winfo_exists() or self.preview_pil_image is None:
            return
        width = max(300, self.popup.winfo_width() - 40)
        height = max(300, self.popup.winfo_height() - 60)
        image = self._build_zoomed_image(
            self.preview_pil_image,
            width,
            height,
            self.popup_zoom,
            self.popup_pan_x,
            self.popup_pan_y,
        )
        self.popup_image = ImageTk.PhotoImage(image)
        if self.popup_label is not None:
            self.popup_label.configure(image=self.popup_image, text="")

    def _build_zoomed_image(self, image: Image.Image, width: int, height: int, zoom: float, pan_x: float, pan_y: float) -> Image.Image:
        fitted = image.copy()
        fitted.thumbnail((width, height), Image.Resampling.LANCZOS)
        if zoom <= 1.001:
            return fitted

        scaled_width = max(1, int(fitted.width * zoom))
        scaled_height = max(1, int(fitted.height * zoom))
        scaled = fitted.resize((scaled_width, scaled_height), Image.Resampling.LANCZOS)
        if scaled_width <= width and scaled_height <= height:
            return scaled

        max_left = max(0, scaled_width - width)
        max_top = max(0, scaled_height - height)
        left = int(min(max(0.0, (scaled_width - width) / 2.0 + pan_x), float(max_left)))
        top = int(min(max(0.0, (scaled_height - height) / 2.0 + pan_y), float(max_top)))
        right = min(scaled_width, left + width)
        bottom = min(scaled_height, top + height)
        return scaled.crop((left, top, right, bottom))

    def _render_main_preview(self):
        if self.preview_pil_image is None:
            return
        width = max(320, self.preview_label.winfo_width() - 20)
        height = max(320, self.preview_label.winfo_height() - 20)
        image = self._build_zoomed_image(
            self.preview_pil_image,
            width,
            height,
            self.preview_zoom,
            self.preview_pan_x,
            self.preview_pan_y,
        )
        self.preview_image = ImageTk.PhotoImage(image)
        self.preview_label.configure(image=self.preview_image, text="")

    def _zoom_preview(self, factor: float):
        if self.preview_pil_image is None:
            return
        self.preview_zoom = min(6.0, max(1.0, self.preview_zoom * factor))
        self._render_main_preview()
        if self.popup is not None and self.popup.winfo_exists():
            self.popup_zoom = min(8.0, max(1.0, self.popup_zoom * factor))
            self._refresh_popup()

    def _reset_zoom(self):
        if self.preview_pil_image is None:
            return
        self.preview_zoom = 1.0
        self.popup_zoom = 1.0
        self.preview_pan_x = 0.0
        self.preview_pan_y = 0.0
        self.popup_pan_x = 0.0
        self.popup_pan_y = 0.0
        self._render_main_preview()
        if self.popup is not None and self.popup.winfo_exists():
            self._refresh_popup()

    def _on_mousewheel(self, event):
        if getattr(event, "delta", 0) > 0 or getattr(event, "num", None) == 4:
            self._zoom_preview(1.15)
        elif getattr(event, "delta", 0) < 0 or getattr(event, "num", None) == 5:
            self._zoom_preview(1 / 1.15)

    def _on_popup_mousewheel(self, event):
        if self.preview_pil_image is None:
            return
        if getattr(event, "delta", 0) > 0 or getattr(event, "num", None) == 4:
            self.popup_zoom = min(8.0, max(1.0, self.popup_zoom * 1.15))
        elif getattr(event, "delta", 0) < 0 or getattr(event, "num", None) == 5:
            self.popup_zoom = min(8.0, max(1.0, self.popup_zoom / 1.15))
        self._refresh_popup()

    def _start_pan(self, target: str, event):
        if self.preview_pil_image is None:
            return
        self._drag_target = target
        self._drag_start_x = event.x
        self._drag_start_y = event.y

    def _drag_pan(self, target: str, event):
        if self.preview_pil_image is None or self._drag_target != target:
            return
        delta_x = event.x - self._drag_start_x
        delta_y = event.y - self._drag_start_y
        self._drag_start_x = event.x
        self._drag_start_y = event.y

        if target == "main" and self.preview_zoom > 1.0:
            self.preview_pan_x -= delta_x
            self.preview_pan_y -= delta_y
            self._render_main_preview()
        elif target == "popup" and self.popup_zoom > 1.0:
            self.popup_pan_x -= delta_x
            self.popup_pan_y -= delta_y
            self._refresh_popup()

    def _end_pan(self):
        self._drag_target = None


def main():
    args = parse_args()
    initial_json_path = None

    if not args.viewer_only:
        if args.mode is not None:
            startup = {
                "run_analysis": True,
                "mode": args.mode,
                "clearance_mm": args.clearance_mm if args.clearance_mm is not None else 5.0,
                "part_number": args.part_number or "",
                "part_number_a": args.part_number_a or "",
                "part_number_b": args.part_number_b or "",
            }
        else:
            startup = prompt_startup_run_options()

        if startup.get("run_analysis"):
            command = build_agent_command(
                startup["mode"],
                startup["clearance_mm"],
                startup["part_number"],
                startup["part_number_a"],
                startup["part_number_b"],
            )
            completed = subprocess.run(command, cwd=str(REPO_ROOT))
            if completed.returncode != 0:
                raise SystemExit(completed.returncode)
            initial_json_path = find_latest_json()
    else:
        initial_json_path = find_latest_json()

    root = tk.Tk()
    DMUPreviewApp(root, initial_json_path=initial_json_path)
    root.mainloop()


if __name__ == "__main__":
    main()
