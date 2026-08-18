from __future__ import annotations

import math
import os
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import threading
from pathlib import Path
import tkinter as tk
from urllib.parse import parse_qs, quote, unquote, urlparse
import webbrowser
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from core.ad import ADCalculator
from core.contracts import ERBASubtype
from core.fingerprint import rdkit_fp_from_smiles
from core.graph import save_all_batch_graphs, save_single_ad_plot
from core.molecule_image import save_molecule_image
from core.predictor import KerasPredictor
from core.pubchem import cas_to_smiles
from core.paths import validate_mutable_directory
from gui.erba_tab import ErbaTab

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


class MainWindow(tk.Tk):
    def __init__(self, project_root: str, output_root: str | None = None, state_root: str | None = None, event_log=None):
        super().__init__()
        self.project_root = project_root
        self.output_root = output_root or os.path.join(self.project_root, "output")
        self.state_root = state_root
        self.event_log = event_log
        self.title("ER Predictor")
        self.configure_initial_window_size()

        self.predictor = KerasPredictor()
        self.ad_calculator = ADCalculator(
            coverage=0.95,
            robust=False,
            empirical=True,
            cache_root=os.path.join(self.state_root, "cache") if self.state_root else None,
        )
        self.last_batch_result = None
        self.last_batch_fp = None
        self.preview_df = None
        self.preview_sort_column = None
        self.preview_sort_ascending = True
        self.graph_paths_by_name = {}
        self.loaded_images = []  # Keep references for Tkinter images.
        self.single_preview_size = (250, 180)
        self.nearest_reference_preview_size = (250, 180)
        self.single_ad_preview_size = (520, 300)
        self.batch_graph_preview_size = (430, 360)
        self.main_thread = threading.current_thread()
        self.options_visible = False

        self.model_path_var = tk.StringVar(value=self.default_model_path())
        self.ad_ref_path_var = tk.StringVar(value=self.default_ad_reference_path())
        self.cas_var = tk.StringVar()
        self.smiles_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Starting...")
        self.prediction_summary_var = tk.StringVar(value="Prediction: -")
        self.active_probability_var = tk.StringVar(value="Probability Positive: -")
        self.inactive_probability_var = tk.StringVar(value="Probability Negative: -")
        self.ad_domain_var = tk.StringVar(value="Applicability domain: Not evaluated")
        self.nearest_reference_var = tk.StringVar(value="Nearest training reference: -")
        self.batch_ad_summary_var = tk.StringVar(value="AD graph will appear after batch prediction.")
        self.batch_input_var = tk.StringVar(value=os.path.join(self.project_root, "templates", "ERTA_KRICT_example.xlsx"))
        self.output_dir_var = tk.StringVar(value=self.output_root)
        default_batch_name = os.path.basename(self.batch_input_var.get()) if os.path.exists(self.batch_input_var.get()) else "No input template selected"
        self.batch_input_display_var = tk.StringVar(value=default_batch_name)
        self.output_dir_display_var = tk.StringVar(value=self.output_dir_var.get())

        self._build_ui()
        self.bind_all("<Alt-e>", lambda _event: self.notebook.select(self.erta_tab))
        self.bind_all("<Alt-a>", lambda _event: self.notebook.select(self.eralpha_tab))
        self.bind_all("<Control-Key-1>", lambda _event: self.notebook.select(self.erta_tab))
        self.bind_all("<Control-Key-2>", lambda _event: self.notebook.select(self.eralpha_tab))
        self.after(300, self.load_defaults_on_startup)

    def configure_initial_window_size(self):
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(1280, max(1180, screen_w - 90))
        height = min(880, max(800, screen_h - 80))
        min_w = min(1160, width)
        min_h = min(780, height)
        x = max((screen_w - width) // 2, 0)
        y = max((screen_h - height) // 2, 0)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.minsize(min_w, min_h)

    def default_model_path(self):
        p = os.path.join(self.project_root, "models", "best_seed831053.keras")
        return p if os.path.exists(p) else ""

    def default_ad_reference_path(self):
        p = os.path.join(self.project_root, "data", "training_reference_RDKit_FP.xlsx")
        return p if os.path.exists(p) else ""

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=8)
        self.erta_tab = ttk.Frame(self.notebook)
        self.eralpha_tab = ErbaTab(
            self.notebook,
            self.project_root,
            output_root=self.output_root,
            event_log=self.event_log,
            structure_editor=self.open_jsme_popup,
            fixed_subtype=ERBASubtype.ER_ALPHA,
        )
        self.erba_tab = self.eralpha_tab
        self.notebook.add(self.erta_tab, text="ERTA")
        self.notebook.add(self.eralpha_tab, text="ERalpha")
        self.notebook.select(self.erta_tab)
        self.notebook.bind("<<NotebookTabChanged>>", self._top_level_tab_changed)

        self.erta_tab.columnconfigure(0, weight=1)
        self.erta_tab.rowconfigure(2, weight=1)
        self._build_header()
        self._build_notebook()
        self._build_status_bar()

    def _top_level_tab_changed(self, _event=None):
        if self.event_log is None:
            return
        selected = self.notebook.select()
        tab = self.notebook.tab(selected, "text") if selected else ""
        self.event_log.emit("workflow.tab_changed", workflow=tab.lower(), tab=tab)

    def _build_header(self):
        summary = ttk.Frame(self.erta_tab)
        summary.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 2))
        summary.columnconfigure(0, weight=1)

        ttk.Button(summary, text="Options", command=self.toggle_options).grid(row=0, column=1, sticky="e")

        frame = ttk.LabelFrame(self.erta_tab, text="Options")
        self.options_frame = frame
        frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(2, 8))
        frame.grid_remove()
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Model (.keras)").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(frame, textvariable=self.model_path_var).grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(frame, text="Browse", command=self.browse_model).grid(row=0, column=2, padx=6, pady=4)
        ttk.Button(frame, text="Reload model", command=self.load_model_clicked).grid(row=0, column=3, padx=6, pady=4)

        ttk.Label(frame, text="AD reference").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(frame, textvariable=self.ad_ref_path_var).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(frame, text="Browse", command=self.browse_ad_reference).grid(row=1, column=2, padx=6, pady=4)
        ttk.Button(frame, text="Reload AD", command=self.fit_ad_clicked).grid(row=1, column=3, padx=6, pady=4)

        note = ttk.Label(frame, text="AD is fitted automatically at startup. Reload AD only when changing the reference file.")
        note.grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 4))

    def _build_notebook(self):
        self.erta_notebook = ttk.Notebook(self.erta_tab)
        self.erta_notebook.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)

        self.single_tab = ttk.Frame(self.erta_notebook)
        self.batch_tab = ttk.Frame(self.erta_notebook)
        self.erta_notebook.add(self.single_tab, text="Single prediction")
        self.erta_notebook.add(self.batch_tab, text="Batch prediction")

        self._build_single_tab()
        self._build_batch_tab()

    def _build_status_bar(self):
        bar = ttk.Label(self.erta_tab, textvariable=self.status_var, anchor="w")
        bar.grid(row=3, column=0, sticky="ew", padx=10, pady=(2, 8))

    def _build_single_tab(self):
        self.single_tab.columnconfigure(0, weight=1)
        self.single_tab.columnconfigure(1, weight=1)
        self.single_tab.rowconfigure(1, weight=1)

        input_frame = ttk.LabelFrame(self.single_tab, text="Input")
        input_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        input_frame.columnconfigure(1, weight=1)

        ttk.Label(input_frame, text="CAS").grid(row=0, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(input_frame, textvariable=self.cas_var).grid(row=0, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(input_frame, text="PubChem search", command=self.pubchem_clicked).grid(row=0, column=2, padx=6, pady=5)

        ttk.Label(input_frame, text="SMILES").grid(row=1, column=0, sticky="w", padx=6, pady=5)
        ttk.Entry(input_frame, textvariable=self.smiles_var).grid(row=1, column=1, sticky="ew", padx=6, pady=5)
        ttk.Button(input_frame, text="Predict", command=self.single_predict_clicked).grid(row=1, column=2, padx=6, pady=5)
        ttk.Button(input_frame, text="Draw structure", command=self.open_jsme_popup).grid(row=1, column=3, padx=6, pady=5)

        result_frame = ttk.LabelFrame(self.single_tab, text="Prediction result")
        result_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(3, weight=1)

        summary_frame = ttk.Frame(result_frame)
        summary_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
        summary_frame.columnconfigure(0, weight=1)

        self.prediction_label = tk.Label(summary_frame, textvariable=self.prediction_summary_var, font=("Segoe UI", 14, "bold"), anchor="w")
        self.prediction_label.grid(row=0, column=0, sticky="w")
        self.ad_domain_label = tk.Label(summary_frame, textvariable=self.ad_domain_var, font=("Segoe UI", 13, "bold"), anchor="w")
        self.ad_domain_label.grid(row=1, column=0, sticky="w", pady=(8, 0))

        self.prob_canvas = tk.Canvas(result_frame, height=92, bg="white", highlightthickness=1, highlightbackground="#d0d7de")
        self.prob_canvas.grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 8))
        self.prob_canvas.bind("<Configure>", lambda event: self.draw_probability_graph())

        ttk.Label(result_frame, text="Details").grid(row=2, column=0, sticky="w", padx=8)
        self.result_text = tk.Text(result_frame, height=7, wrap="word", font=("Consolas", 9))
        self.result_text.tag_configure("detail_key", font=("Consolas", 9, "bold"))
        self.result_text.tag_configure("detail_value", font=("Consolas", 9))
        self.result_text.grid(row=3, column=0, sticky="nsew", padx=8, pady=(2, 8))

        right_frame = ttk.Frame(self.single_tab)
        right_frame.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=0)
        right_frame.rowconfigure(1, weight=1)

        molecule_pair_frame = ttk.Frame(right_frame)
        molecule_pair_frame.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        molecule_pair_frame.columnconfigure(0, weight=1)
        molecule_pair_frame.columnconfigure(1, weight=1)

        struct_frame = ttk.LabelFrame(molecule_pair_frame, text="Input molecule")
        struct_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.structure_label = ttk.Label(struct_frame, text="No structure")
        self.structure_label.pack(expand=True, fill="both", padx=6, pady=6)

        ref_frame = ttk.LabelFrame(molecule_pair_frame, text="Nearest training reference")
        ref_frame.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        self.nearest_reference_structure_label = ttk.Label(ref_frame, text="No reference")
        self.nearest_reference_structure_label.pack(expand=True, fill="both", padx=6, pady=6)

        single_ad_frame = ttk.LabelFrame(right_frame, text="Applicability domain")
        single_ad_frame.grid(row=1, column=0, sticky="nsew", pady=(6, 0))
        single_ad_frame.columnconfigure(0, weight=1)
        single_ad_frame.rowconfigure(0, weight=1)
        self.single_ad_graph_label = ttk.Label(single_ad_frame, text="AD graph will appear after prediction.")
        self.single_ad_graph_label.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.nearest_reference_label = ttk.Label(single_ad_frame, textvariable=self.nearest_reference_var, wraplength=420, justify="left")
        self.nearest_reference_label.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))

        note = ttk.Label(self.single_tab, text="CAS to SMILES requires internet access. Direct SMILES prediction works offline.")
        note.grid(row=2, column=0, columnspan=2, sticky="sw", padx=8, pady=8)

    def _build_batch_tab(self):
        self.batch_tab.columnconfigure(0, weight=3, uniform="batch")
        self.batch_tab.columnconfigure(1, weight=2, uniform="batch")
        self.batch_tab.rowconfigure(1, weight=1)

        frame = ttk.LabelFrame(self.batch_tab, text="Batch")
        frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=8, pady=8)
        frame.columnconfigure(1, weight=1)
        frame.columnconfigure(3, weight=1)

        ttk.Button(frame, text="Download template", command=self.download_template_clicked).grid(row=0, column=0, padx=6, pady=5)

        ttk.Button(frame, text="Input template", command=self.browse_batch_input).grid(row=0, column=1, sticky="e", padx=6, pady=5)
        ttk.Label(frame, textvariable=self.batch_input_display_var, anchor="w").grid(row=0, column=2, sticky="ew", padx=6, pady=5)

        ttk.Button(frame, text="Output directory", command=self.browse_output_dir).grid(row=1, column=1, sticky="e", padx=6, pady=5)
        ttk.Label(frame, textvariable=self.output_dir_display_var, anchor="w").grid(row=1, column=2, sticky="ew", padx=6, pady=5)
        self.run_batch_button = tk.Button(
            frame,
            text="Run batch prediction",
            command=self.batch_predict_clicked,
            font=("Segoe UI", 10, "bold"),
            bg="#0969da",
            fg="white",
            activebackground="#0757b5",
            activeforeground="white",
            relief="raised",
            bd=1,
            cursor="hand2",
        )
        self.run_batch_button.grid(row=1, column=3, sticky="ew", padx=8, pady=6, ipadx=18, ipady=5)

        table_frame = ttk.LabelFrame(self.batch_tab, text="Prediction result")
        table_frame.grid(row=1, column=0, sticky="nsew", padx=8, pady=8)
        table_frame.configure(width=650)
        table_frame.grid_propagate(False)
        table_frame.rowconfigure(1, weight=1)
        table_frame.columnconfigure(0, weight=1)
        self.batch_summary_text = tk.Text(table_frame, height=6, wrap="word", font=("Consolas", 9))
        self.batch_summary_text.tag_configure("detail_key", font=("Consolas", 9, "bold"))
        self.batch_summary_text.tag_configure("detail_value", font=("Consolas", 9))
        self.batch_summary_text.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 4))
        self.batch_summary_text.insert(tk.END, "Run batch prediction to show summary.")
        self.batch_summary_text.configure(state="disabled")
        self.batch_tree = ttk.Treeview(table_frame, show="headings")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.batch_tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.batch_tree.xview)
        self.batch_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.batch_tree.grid(row=1, column=0, sticky="nsew")
        self.batch_tree.bind("<Double-1>", self.on_batch_tree_double_click)
        yscroll.grid(row=1, column=1, sticky="ns")
        xscroll.grid(row=2, column=0, sticky="ew")

        graph_frame = ttk.LabelFrame(self.batch_tab, text="Applicability domain")
        graph_frame.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        graph_frame.configure(width=450)
        graph_frame.grid_propagate(False)
        graph_frame.columnconfigure(0, weight=1)
        graph_frame.rowconfigure(1, weight=1)
        graph_top = ttk.Frame(graph_frame)
        graph_top.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ttk.Button(graph_top, text="Refresh", command=self.refresh_graph_list).pack(side="left", padx=(0, 4))
        self.graph_combo = ttk.Combobox(graph_top, values=[], state="readonly", width=42)
        self.graph_combo.pack(side="left", fill="x", expand=True)
        self.graph_combo.bind("<<ComboboxSelected>>", lambda e: self.display_selected_graph())
        self.graph_label = ttk.Label(graph_frame, text="Graphs will appear after batch prediction.")
        self.graph_label.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        self.batch_ad_summary_label = ttk.Label(graph_frame, textvariable=self.batch_ad_summary_var, wraplength=430, justify="left")
        self.batch_ad_summary_label.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))

    # ---------- UI helpers ----------
    def ui(self, func, *args, **kwargs):
        if threading.current_thread() is self.main_thread:
            return func(*args, **kwargs)
        self.after(0, lambda: func(*args, **kwargs))

    def toggle_options(self):
        self.options_visible = not self.options_visible
        if self.options_visible:
            self.options_frame.grid()
        else:
            self.options_frame.grid_remove()

    def set_status(self, text: str):
        if threading.current_thread() is not self.main_thread:
            self.after(0, lambda: self.set_status(text))
            return
        self.status_var.set(text)
        self.update_idletasks()

    def run_threaded(self, func):
        t = threading.Thread(target=func, daemon=True)
        t.start()

    def show_error(self, title, err):
        if threading.current_thread() is not self.main_thread:
            self.after(0, lambda: self.show_error(title, err))
            return
        messagebox.showerror(title, str(err))
        self.set_status("Error")

    def draw_probability_graph(self, prob_active=None):
        if prob_active is None:
            prob_active = getattr(self, "_last_prob_active", None)

        canvas = self.prob_canvas
        canvas.delete("all")
        width = max(canvas.winfo_width(), 320)
        margin = 18
        label_width = 72
        bar_x = margin + label_width
        bar_w = max(width - bar_x - margin - 48, 120)
        inactive_y = 22
        active_y = 58
        bar_h = 18

        def draw_bar(y, label, value, color):
            value = max(0.0, min(1.0, float(value)))
            canvas.create_text(margin, y + bar_h / 2, text=label, anchor="w", fill="#24292f")
            canvas.create_rectangle(bar_x, y, bar_x + bar_w, y + bar_h, fill="#eef2f7", outline="#d0d7de")
            canvas.create_rectangle(bar_x, y, bar_x + bar_w * value, y + bar_h, fill=color, outline=color)
            canvas.create_text(bar_x + bar_w + 8, y + bar_h / 2, text=f"{value:.3f}", anchor="w", fill="#24292f")

        if prob_active is None:
            draw_bar(inactive_y, "Negative", 0.0, "#2da44e")
            draw_bar(active_y, "Positive", 0.0, "#cf222e")
            canvas.create_text(bar_x, 8, text="Run prediction to show probabilities", anchor="w", fill="#57606a")
            return

        prob_active = max(0.0, min(1.0, float(prob_active)))
        prob_inactive = 1.0 - prob_active
        draw_bar(inactive_y, "Negative", prob_inactive, "#2da44e")
        draw_bar(active_y, "Positive", prob_active, "#cf222e")

    @staticmethod
    def display_label_value(key, value):
        text = str(value).strip()
        if str(key).strip().lower() in {"label", "activity", "active", "inactive"}:
            if text in {"1", "1.0"}:
                return "Positive"
            if text in {"0", "0.0"}:
                return "Negative"
        return text

    @staticmethod
    def format_nearest_reference(ad_res) -> str:
        if ad_res is None or not ad_res.fitted:
            return "Nearest training reference: -"
        ref = getattr(ad_res, "display_nearest_reference", None) or ad_res.nearest_reference or {}
        display_similarity = getattr(ad_res, "display_similarity", None)
        if display_similarity is None:
            display_similarity = ad_res.max_similarity
        parts = []
        for key in ["Reference row", "CAS", "CAS No", "CAS No.", "CAS RN", "CID", "PubChem_CID", "Name", "Chemical name", "SMILES", "Canonical_SMILES", "label", "Activity"]:
            value = ref.get(key)
            if value:
                text = MainWindow.display_label_value(key, value)
                if key.upper().endswith("SMILES") or key == "SMILES":
                    text = text[:80]
                parts.append(f"{key}: {text}")
            if len(parts) >= 4:
                break
        if not parts:
            for key, value in ref.items():
                if value:
                    text = MainWindow.display_label_value(key, value)
                    if "SMILES" in str(key).upper():
                        text = text[:80]
                    parts.append(f"{key}: {text}")
                if len(parts) >= 4:
                    break
        display_index = getattr(ad_res, "display_nearest_index", None)
        suffix = " / ".join(parts) if parts else f"Reference index: {display_index}"
        return f"Nearest training reference: Morgan similarity={display_similarity:.4f} / {suffix}"

    @staticmethod
    def nearest_reference_smiles_and_label(ad_res):
        if ad_res is None or not ad_res.fitted:
            return "", ""
        ref = getattr(ad_res, "display_nearest_reference", None) or ad_res.nearest_reference or {}
        smiles = ""
        for key in ["SMILES", "Canonical_SMILES", "Canonical SMILES", "Isomeric_SMILES", "Isomeric SMILES"]:
            if ref.get(key):
                smiles = str(ref.get(key)).strip()
                break
        label_parts = []
        for key in ["label", "Label", "Activity", "active", "inactive"]:
            if ref.get(key):
                label_parts.append(f"label: {MainWindow.display_label_value(key, ref.get(key))}")
                break
        for key in ["CID", "PubChem_CID"]:
            if ref.get(key):
                label_parts.append(f"CID: {ref.get(key)}")
                break
        return smiles, " / ".join(label_parts)

    def render_single_result(self, result, ad_res, detail_lines, ad_status: str | None = None, ad_graph_path: str | None = None):
        self._last_prob_active = result.probability_active
        prediction_display = "Positive" if result.prediction == 1 else "Negative"
        prediction_color = "#cf222e" if result.prediction == 1 else "#1a7f37"
        self.prediction_summary_var.set(f"Prediction: {prediction_display}")
        self.prediction_label.configure(fg=prediction_color)
        self.active_probability_var.set(f"Probability Positive: {result.probability_active:.6f}")
        self.inactive_probability_var.set(f"Probability Negative: {result.probability_inactive:.6f}")

        if ad_res is not None and ad_res.fitted:
            domain = "In-domain" if ad_res.in_domain else "Out-of-domain"
            self.ad_domain_var.set(f"Applicability domain: {domain}")
            self.ad_domain_label.configure(fg="#1a7f37" if ad_res.in_domain else "#cf222e")
            self.nearest_reference_var.set(self.format_nearest_reference(ad_res))
            self.display_nearest_reference_structure(ad_res)
        elif ad_status:
            self.ad_domain_var.set(f"Applicability domain: {ad_status}")
            self.ad_domain_label.configure(fg="#57606a")
            self.nearest_reference_var.set("Nearest training reference: -")
            self.clear_nearest_reference_structure()
        else:
            self.ad_domain_var.set("Applicability domain: Not evaluated")
            self.ad_domain_label.configure(fg="#57606a")
            self.nearest_reference_var.set("Nearest training reference: -")
            self.clear_nearest_reference_structure()

        self.draw_probability_graph(result.probability_active)
        self.render_detail_lines(detail_lines)
        self.draw_structure(result.canonical_smiles or result.smiles)
        if ad_graph_path:
            self.display_single_ad_graph(ad_graph_path)
        elif ad_status:
            self.clear_single_ad_graph(f"AD graph unavailable: {ad_status}")

    def render_detail_lines(self, detail_lines):
        self.result_text.delete("1.0", tk.END)
        for line in detail_lines:
            text = str(line)
            if ":" in text:
                key, value = text.split(":", 1)
                self.result_text.insert(tk.END, f"{key}:", "detail_key")
                self.result_text.insert(tk.END, f"{value}\n", "detail_value")
            else:
                self.result_text.insert(tk.END, f"{text}\n", "detail_value")

    def render_text_lines(self, widget, detail_lines):
        widget.configure(state="normal")
        widget.delete("1.0", tk.END)
        for line in detail_lines:
            text = str(line)
            if ":" in text:
                key, value = text.split(":", 1)
                widget.insert(tk.END, f"{key}:", "detail_key")
                widget.insert(tk.END, f"{value}\n", "detail_value")
            else:
                widget.insert(tk.END, f"{text}\n", "detail_value")
        widget.configure(state="disabled")

    def ensure_ad_fitted(self, reference_path: str, force_refit: bool = False):
        ensure = getattr(self.ad_calculator, "ensure_fitted_from_excel_cached", None)
        if ensure is not None:
            return (
                ensure(reference_path, force_refit=True)
                if force_refit
                else ensure(reference_path)
            )
        return (
            self.ad_calculator.fit_from_excel_cached(reference_path, force_refit=True)
            if force_refit
            else self.ad_calculator.fit_from_excel_cached(reference_path)
        )

    def load_defaults_on_startup(self):
        def job():
            failures = []

            model_path = self.model_path_var.get()
            if not model_path.strip():
                failures.append("model path is blank")
            elif not os.path.exists(model_path):
                failures.append(f"model path does not exist: {model_path}")
            else:
                try:
                    self.set_status("Loading default model...")
                    self.predictor.load_model(model_path)
                except Exception as e:
                    failures.append(f"model load failed: {e}")

            ad_path = self.ad_ref_path_var.get()
            if not ad_path.strip():
                failures.append("AD reference path is blank")
            elif not os.path.exists(ad_path):
                failures.append(f"AD reference path does not exist: {ad_path}")
            else:
                try:
                    self.set_status("Loading default AD reference...")
                    self.ensure_ad_fitted(ad_path)
                except Exception as e:
                    failures.append(f"AD fit failed: {e}")

            self.set_status("Ready" if not failures else "Startup failed: " + " | ".join(failures))

        self.run_threaded(job)

    def browse_model(self):
        path = filedialog.askopenfilename(title="Select Keras model", filetypes=[("Keras model", "*.keras *.h5"), ("All files", "*.*")])
        if path:
            self.model_path_var.set(path)

    def browse_ad_reference(self):
        path = filedialog.askopenfilename(title="Select AD reference Excel", filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.ad_ref_path_var.set(path)

    def browse_batch_input(self):
        path = filedialog.askopenfilename(title="Select input Excel", filetypes=[("Excel", "*.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.batch_input_var.set(path)
            self.batch_input_display_var.set(os.path.basename(path))

    def browse_output_dir(self):
        current = self.output_dir_var.get()
        initialdir = current if current and os.path.exists(current) else self.output_root
        path = filedialog.askdirectory(title="Select output directory", initialdir=initialdir)
        if not path:
            return
        try:
            validated = validate_mutable_directory(path, forbidden_roots=(self.project_root,))
        except Exception as error:
            self.show_error("Invalid output directory", error)
            return
        self.output_dir_var.set(str(validated))
        self.output_dir_display_var.set(str(validated))

    def validated_output_dir(self) -> str:
        return str(validate_mutable_directory(
            self.output_dir_var.get(),
            forbidden_roots=(self.project_root,),
        ))

    @staticmethod
    def _is_blank(value) -> bool:
        return pd.isna(value) or str(value).strip() == ""

    @staticmethod
    def _find_column(df: pd.DataFrame, candidates):
        normalized = {str(col).strip().lower(): col for col in df.columns}
        for candidate in candidates:
            col = normalized.get(candidate.strip().lower())
            if col is not None:
                return col
        return None

    def prepare_batch_input(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(col).strip() for col in df.columns]

        cas_col = self._find_column(df, ["CAS", "CAS No", "CAS No.", "CAS RN", "CASRN", "CAS Number", "CAS_Number"])
        smiles_col = self._find_column(df, ["SMILES", "Canonical_SMILES", "Canonical SMILES", "Isomeric_SMILES", "Isomeric SMILES"])

        if smiles_col is None:
            df["SMILES"] = ""
            smiles_col = "SMILES"
        elif smiles_col != "SMILES":
            df["SMILES"] = df[smiles_col]
            smiles_col = "SMILES"

        if cas_col is not None and cas_col != "CAS":
            df["CAS"] = df[cas_col]
            cas_col = "CAS"

        if cas_col is None:
            return df

        total = len(df)
        for idx, row in df.iterrows():
            current_smiles = row.get(smiles_col, "")
            if not self._is_blank(current_smiles):
                continue

            cas = row.get(cas_col, "")
            if self._is_blank(cas):
                df.at[idx, "PubChem_status"] = "Skipped: CAS is empty"
                continue

            cas_text = str(cas).strip()
            try:
                self.set_status(f"Fetching SMILES from PubChem: {idx + 1} / {total} ({cas_text})")
                res = cas_to_smiles(cas_text)
                smiles = res.get("CanonicalSMILES") or res.get("IsomericSMILES")
                if not smiles:
                    raise RuntimeError("PubChem did not return a SMILES string.")
                df.at[idx, "SMILES"] = smiles
                df.at[idx, "PubChem_CID"] = res.get("PubChem_CID", "")
                df.at[idx, "PubChem_status"] = "Found"
            except Exception as err:
                df.at[idx, "PubChem_status"] = f"Not found: {err}"

        return df

    # ---------- Actions ----------
    def load_model_clicked(self):
        def job():
            try:
                self.set_status("Loading model...")
                self.predictor.load_model(self.model_path_var.get())
                self.set_status(f"Model loaded: {self.predictor.model_name}, input shape={self.predictor.input_shape}")
                self.ui(messagebox.showinfo, "Model loaded", f"Loaded model:\n{self.predictor.model_path}\n\nInput shape: {self.predictor.input_shape}")
            except Exception as e:
                self.show_error("Model load failed", e)
        self.run_threaded(job)

    def fit_ad_clicked(self):
        def job():
            try:
                self.set_status("Rebuilding AD reference cache...")
                self.ensure_ad_fitted(
                    self.ad_ref_path_var.get(),
                    force_refit=True,
                )
                cache_status = self.ad_calculator.last_cache_status or "AD cache status unavailable."
                self.set_status(f"AD fitted. kNN 95% mean-distance threshold={self.ad_calculator.threshold:.4f}, Similarity cutoff={self.ad_calculator.similarity_threshold:.2f}. {cache_status}")
                self.ui(messagebox.showinfo, "AD fitted", f"AD reference fitted.\nkNN 95% mean-distance threshold: {self.ad_calculator.threshold:.4f}\nSimilarity cutoff: {self.ad_calculator.similarity_threshold:.2f}\n\n{cache_status}")
            except Exception as e:
                self.show_error("AD fitting failed", e)
        self.run_threaded(job)

    def pubchem_clicked(self):
        def job():
            try:
                self.set_status("Searching PubChem...")
                res = cas_to_smiles(self.cas_var.get())
                smiles = res.get("CanonicalSMILES") or res.get("IsomericSMILES")
                if not smiles:
                    raise RuntimeError("PubChem did not return a SMILES string.")
                self.ui(self.smiles_var.set, smiles)
                self.set_status(f"PubChem found CID {res.get('PubChem_CID')}")
            except Exception as e:
                self.show_error("PubChem search failed", e)
        self.run_threaded(job)

    def open_structure_editor(self):
        StructureEditorDialog(self, self.smiles_var)

    def open_jsme_popup(self, smiles_var=None, status_var=None):
        target_smiles_var = smiles_var or self.smiles_var
        try:
            server = self.start_jsme_server(target_smiles_var, status_var)
            url = f"http://127.0.0.1:{server.server_port}/"
            webbrowser.open(url, new=1)
            message = "JSME editor opened. Draw a structure, then click Save SMILES."
            status_var.set(message) if status_var is not None else self.set_status(message)
        except Exception as e:
            self.show_error("JSME editor failed", e)

    def start_jsme_server(self, smiles_var=None, status_var=None):
        app = self
        target_smiles_var = smiles_var or self.smiles_var
        server_box = {}

        class JSMEHandler(BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urlparse(self.path)
                if parsed.path == "/save":
                    params = parse_qs(parsed.query)
                    smiles = unquote(params.get("smiles", [""])[0]).strip()
                    if smiles:
                        app.after(
                            0,
                            lambda: app.import_drawn_smiles(
                                smiles, target_smiles_var, status_var
                            ),
                        )
                    self.send_html(app.jsme_saved_html(smiles))
                    threading.Thread(target=server_box["server"].shutdown, daemon=True).start()
                    return
                self.send_html(app.jsme_editor_html(target_smiles_var.get().strip()))

            def send_html(self, html: str):
                data = html.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def log_message(self, format, *args):
                return

        server = ThreadingHTTPServer(("127.0.0.1", 0), JSMEHandler)
        server_box["server"] = server
        threading.Thread(target=server.serve_forever, daemon=True).start()
        return server

    def import_drawn_smiles(self, smiles: str, smiles_var=None, status_var=None):
        (smiles_var or self.smiles_var).set(smiles)
        message = f"JSME SMILES imported: {smiles}"
        status_var.set(message) if status_var is not None else self.set_status(message)

    @staticmethod
    def jsme_editor_html(initial_smiles: str) -> str:
        initial = quote(initial_smiles or "")
        return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Draw structure - JSME</title>
  <style>
    body {{
      margin: 0;
      font-family: Segoe UI, Arial, sans-serif;
      background: #f6f8fa;
      color: #24292f;
    }}
    .topbar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 10px 14px;
      border-bottom: 1px solid #d0d7de;
      background: white;
    }}
    .title {{
      font-weight: 700;
    }}
    .hint {{
      font-size: 12px;
      color: #57606a;
    }}
    #jsme_container {{
      width: calc(100vw - 24px);
      height: calc(100vh - 104px);
      margin: 12px;
      background: white;
      border: 1px solid #d0d7de;
    }}
    .footer {{
      display: flex;
      gap: 8px;
      align-items: center;
      justify-content: flex-end;
      padding: 0 12px 12px;
    }}
    button {{
      border: 1px solid #d0d7de;
      border-radius: 6px;
      background: white;
      padding: 8px 12px;
      cursor: pointer;
      font-weight: 600;
    }}
    .primary {{
      border-color: #0969da;
      background: #0969da;
      color: white;
    }}
  </style>
  <script>
    var jsmeApplet = null;
    var initialSmiles = decodeURIComponent("{initial}");

    function jsmeOnLoad() {{
      jsmeApplet = new JSApplet.JSME("jsme_container", "100%", "100%", {{
        "options": "oldlook,star"
      }});
      if (initialSmiles) {{
        try {{
          jsmeApplet.readGenericMolecularInput(initialSmiles);
        }} catch (e) {{
          console.warn(e);
        }}
      }}
    }}

    function saveSmiles() {{
      if (!jsmeApplet) {{
        alert("JSME is not ready yet.");
        return;
      }}
      var smiles = jsmeApplet.smiles();
      if (!smiles) {{
        alert("Draw a structure before saving.");
        return;
      }}
      window.location.href = "/save?smiles=" + encodeURIComponent(smiles);
    }}
  </script>
  <script type="text/javascript" src="https://jsme-editor.github.io/dist/jsme/jsme.nocache.js"></script>
</head>
<body>
  <div class="topbar">
    <div>
      <div class="title">JSME Structure Editor</div>
      <div class="hint">Draw a molecule, then click Save SMILES to send it back to ERTA Predictor.</div>
    </div>
    <button class="primary" onclick="saveSmiles()">Save SMILES</button>
  </div>
  <div id="jsme_container">
    <div style="padding:18px">Loading JSME editor...</div>
  </div>
  <div class="footer">
    <button onclick="window.close()">Close</button>
    <button class="primary" onclick="saveSmiles()">Save SMILES</button>
  </div>
</body>
</html>"""

    @staticmethod
    def jsme_saved_html(smiles: str) -> str:
        escaped = smiles.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return f"""<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>SMILES saved</title>
  <style>
    body {{ font-family: Segoe UI, Arial, sans-serif; padding: 24px; color: #24292f; }}
    code {{ display: block; padding: 10px; background: #f6f8fa; border: 1px solid #d0d7de; }}
  </style>
</head>
<body>
  <h3>SMILES saved to ERTA Predictor.</h3>
  <code>{escaped}</code>
  <p>You can close this browser window.</p>
</body>
</html>"""

    def single_predict_clicked(self):
        def job():
            try:
                if not self.predictor.is_loaded():
                    self.predictor.load_model(self.model_path_var.get())
                smiles = self.smiles_var.get().strip()
                if not smiles:
                    raise ValueError("SMILES is empty. Enter SMILES or search PubChem by CAS first.")
                self.set_status("Predicting single molecule...")
                result = self.predictor.predict_smiles(smiles)
                fp, _, _ = rdkit_fp_from_smiles(smiles)

                base_lines = [
                    f"CAS: {self.cas_var.get().strip()}",
                    f"Input SMILES: {result.smiles}",
                    f"Canonical SMILES: {result.canonical_smiles}",
                    f"Mol valid: {result.mol_valid}",
                    f"Negative probability: {result.probability_inactive:.6f}",
                    f"Positive probability: {result.probability_active:.6f}",
                    f"Prediction result: {'Positive' if result.prediction == 1 else 'Negative'}",
                ]

                self.ui(self.render_single_result, result, None, base_lines + ["Applicability domain: Evaluating..."], "Evaluating...")

                lines = list(base_lines)
                ad_res = None
                ad_status = None
                ad_graph_path = None
                try:
                    ad_path = self.ad_ref_path_var.get().strip()
                    if not ad_path:
                        raise ValueError("AD reference path is blank.")
                    if not os.path.exists(ad_path):
                        raise FileNotFoundError(f"AD reference path does not exist: {ad_path}")
                    self.set_status("Loading AD reference...")
                    self.ensure_ad_fitted(ad_path)
                    self.set_status("Evaluating applicability domain...")
                    ad_res = self.ad_calculator.predict(
                        fp.reshape(1, -1),
                        input_smiles=result.canonical_smiles or smiles,
                    )
                    if not ad_res.fitted:
                        raise RuntimeError(ad_res.message or "AD reference is not fitted.")

                    domain = "In-domain" if ad_res.in_domain else "Out-of-domain"
                    lines.extend([
                        f"Applicability domain: {domain}",
                        f"kNN mean distance: {ad_res.distance:.4f} / 95% threshold {ad_res.threshold:.4f}",
                        f"Distance AD: {'In-domain' if ad_res.distance_in_domain else 'Out-of-domain'}",
                        f"AD RDKit similarity: {ad_res.max_similarity:.4f} / cutoff {ad_res.similarity_threshold:.4f}",
                        f"Similarity AD: {'In-domain' if ad_res.similarity_in_domain else 'Out-of-domain'}",
                        self.format_nearest_reference(ad_res),
                        f"PCA position: PC1={ad_res.pc1:.4f}, PC2={ad_res.pc2:.4f}",
                    ])
                except Exception as ad_error:
                    ad_status = f"Unavailable: {ad_error}"
                    lines.append(f"Applicability domain: {ad_status}")

                if ad_res is not None and ad_res.fitted:
                    try:
                        ad_graph_path = save_single_ad_plot(
                            self.ad_calculator,
                            fp.reshape(1, -1),
                            os.path.join(self.validated_output_dir(), "graphs"),
                        )
                    except Exception as graph_error:
                        lines.append(f"AD graph: Unavailable: {graph_error}")

                self.ui(self.render_single_result, result, ad_res, lines, ad_status, ad_graph_path)
                if ad_status:
                    self.set_status(f"Single prediction done: {result.prediction_label}; AD {ad_status}")
                else:
                    self.set_status(f"Single prediction done: {result.prediction_label}")
            except Exception as e:
                self.show_error("Prediction failed", e)
        self.run_threaded(job)

    def draw_structure(self, smiles: str):
        if not PIL_AVAILABLE:
            self.structure_label.configure(text="Pillow is not installed; structure preview unavailable.", image="")
            return
        try:
            out_dir = os.path.join(self.validated_output_dir(), "structures")
            path = os.path.join(out_dir, "single_structure.png")
            save_molecule_image(smiles, path)
            photo = self.make_preview_photo(path, self.single_preview_size)
            self.loaded_images.append(photo)
            self.structure_label.configure(image=photo, text="")
        except Exception as e:
            self.structure_label.configure(text=f"Structure drawing failed:\n{e}", image="")

    def clear_nearest_reference_structure(self):
        self.nearest_reference_structure_label.configure(text="No reference", image="")

    def display_nearest_reference_structure(self, ad_res):
        if not PIL_AVAILABLE:
            self.nearest_reference_structure_label.configure(text="Pillow is not installed.", image="")
            return
        smiles, label = self.nearest_reference_smiles_and_label(ad_res)
        if not smiles:
            self.clear_nearest_reference_structure()
            return
        try:
            out_dir = os.path.join(self.validated_output_dir(), "structures")
            path = os.path.join(out_dir, "nearest_training_reference.png")
            save_molecule_image(smiles, path, size=(260, 200))
            photo = self.make_preview_photo(path, self.nearest_reference_preview_size)
            self.loaded_images.append(photo)
            self.nearest_reference_structure_label.configure(image=photo, text=label or "Nearest reference", compound="top", wraplength=180)
        except Exception as e:
            self.nearest_reference_structure_label.configure(text=f"Reference drawing failed:\n{e}", image="")

    def clear_single_ad_graph(self, message: str = "AD graph will appear after prediction."):
        self.single_ad_graph_label.configure(text=message, image="")

    def display_single_ad_graph(self, path: str):
        if not PIL_AVAILABLE:
            self.single_ad_graph_label.configure(text="Pillow is not installed; AD graph preview unavailable.", image="")
            return
        if not path or not os.path.exists(path):
            self.clear_single_ad_graph("AD graph was not generated.")
            return
        try:
            photo = self.make_preview_photo(path, self.single_ad_preview_size)
            self.loaded_images.append(photo)
            self.single_ad_graph_label.configure(image=photo, text="")
        except Exception as e:
            self.single_ad_graph_label.configure(text=f"AD graph preview failed:\n{e}", image="")

    @staticmethod
    def make_preview_photo(path: str, size: tuple[int, int]):
        img = Image.open(path).convert("RGB")
        img.thumbnail(size)
        canvas = Image.new("RGB", size, "white")
        x = (size[0] - img.width) // 2
        y = (size[1] - img.height) // 2
        canvas.paste(img, (x, y))
        return ImageTk.PhotoImage(canvas)

    def download_template_clicked(self):
        try:
            path = filedialog.asksaveasfilename(
                title="Download batch template",
                defaultextension=".xlsx",
                initialfile="Template.xlsx",
                filetypes=[("Excel", "*.xlsx")],
            )
            if not path:
                return
            validate_mutable_directory(Path(path).parent, forbidden_roots=(self.project_root,))
            df = pd.DataFrame({"CAS": []})
            df.to_excel(path, index=False)
            self.batch_input_var.set(path)
            self.batch_input_display_var.set(os.path.basename(path))
            self.set_status(f"Template downloaded: {path}")
            self.ui(messagebox.showinfo, "Template downloaded", f"Saved:\n{path}\n\nEnter CAS values, save the file, then run batch prediction.")
        except Exception as e:
            self.show_error("Template download failed", e)

    def batch_predict_clicked(self):
        def job():
            try:
                if not self.predictor.is_loaded():
                    self.predictor.load_model(self.model_path_var.get())
                input_path = self.batch_input_var.get()
                if not os.path.exists(input_path):
                    raise FileNotFoundError(input_path)
                out_dir = self.validated_output_dir()
                self.set_status("Reading batch input...")
                df = pd.read_excel(input_path)

                df = self.prepare_batch_input(df)

                self.set_status("Running batch prediction...")
                result_df, fp_df = self.predictor.predict_dataframe(df, smiles_col="SMILES")
                if not self.ad_calculator.fitted and self.ad_ref_path_var.get() and os.path.exists(self.ad_ref_path_var.get()):
                    self.set_status("Loading AD reference...")
                    self.ensure_ad_fitted(self.ad_ref_path_var.get())
                if self.ad_calculator.fitted:
                    z, mean_distance, in_domain, max_similarity, distance_in_domain, similarity_in_domain = self.ad_calculator.transform_with_similarity(fp_df.values.astype(float))
                    result_df["AD"] = ["In-domain" if x else "Out-of-domain" for x in in_domain]
                    result_df["AD_MeanDistance"] = mean_distance
                    result_df["AD_DistanceThreshold"] = self.ad_calculator.threshold
                    result_df["AD_Distance_InDomain"] = distance_in_domain
                    result_df["AD_SimilarityMax"] = max_similarity
                    result_df["AD_SimilarityThreshold"] = self.ad_calculator.similarity_threshold
                    result_df["AD_Similarity_InDomain"] = similarity_in_domain
                    result_df["AD_PC1"] = z[:, 0]
                    result_df["AD_PC2"] = z[:, 1]
                # Preserve PubChem status if generated.
                for col in ["PubChem_CID", "PubChem_status"]:
                    if col in df.columns and col not in result_df.columns:
                        result_df[col] = df[col]

                model_name = self.predictor.model_name or "model"
                base = os.path.splitext(os.path.basename(input_path))[0]
                out_path = os.path.join(out_dir, f"{base}_{model_name}_prediction.xlsx")
                result_df.to_excel(out_path, index=False)
                self.last_batch_result = result_df
                self.last_batch_fp = fp_df
                self.ui(self.update_preview_table, result_df)
                self.ui(self.update_batch_result_summary, result_df, out_path)
                self.ui(self.update_batch_ad_summary, result_df)
                graph_paths = self.generate_batch_graphs(show_errors=True)
                if graph_paths:
                    self.set_status(f"Batch prediction saved: {out_path} / Graphs: {len(graph_paths)}")
                    message = f"Saved:\n{out_path}\n\nGraphs generated:\n{os.path.join(out_dir, 'graphs')}"
                else:
                    self.set_status(f"Batch prediction saved: {out_path} / No graphs generated")
                    message = f"Saved:\n{out_path}\n\nGraphs were not generated. Check the status/error message."
                self.ui(messagebox.showinfo, "Batch prediction done", message)
            except Exception as e:
                self.show_error("Batch prediction failed", e)
        self.run_threaded(job)

    def update_preview_table(self, df: pd.DataFrame, max_rows: int = 200):
        self.preview_df = df.copy()
        self.preview_sort_column = None
        self.preview_sort_ascending = True
        self.render_preview_table(self.preview_df, max_rows=max_rows)

    def update_batch_result_summary(self, df: pd.DataFrame, out_path: str):
        total = int(len(df))
        pred_counts = df["Prediction_label"].value_counts() if "Prediction_label" in df.columns else pd.Series(dtype=int)
        lines = [
            f"Total chemicals: {total}",
            f"Positive: {int(pred_counts.get('Positive', 0))}",
            f"Negative: {int(pred_counts.get('Negative', 0))}",
        ]
        if "AD" in df.columns:
            ad_counts = df["AD"].value_counts()
            lines.extend([
                f"AD In-domain: {int(ad_counts.get('In-domain', 0))}",
                f"AD Out-of-domain: {int(ad_counts.get('Out-of-domain', 0))}",
            ])
        lines.append(f"Output file: {out_path}")
        self.render_text_lines(self.batch_summary_text, lines)

    def update_batch_ad_summary(self, df: pd.DataFrame):
        if "AD" not in df.columns:
            self.batch_ad_summary_var.set("Applicability domain was not evaluated.")
            return
        counts = df["AD"].value_counts()
        in_count = int(counts.get("In-domain", 0))
        out_count = int(counts.get("Out-of-domain", 0))
        total = int(len(df))
        self.batch_ad_summary_var.set(f"Applicability domain: In-domain {in_count} / Out-of-domain {out_count} / Total {total}")

    def render_preview_table(self, df: pd.DataFrame, max_rows: int = 200):
        self.batch_tree.delete(*self.batch_tree.get_children())
        cols = list(df.columns[:12])
        self.batch_tree["columns"] = cols
        for col in cols:
            self.batch_tree.heading(col, text=col)
            self.batch_tree.column(col, width=105, minwidth=80, anchor="w", stretch=False)
        for _, row in df.head(max_rows).iterrows():
            values = []
            for col in cols:
                v = row[col]
                if isinstance(v, float):
                    v = f"{v:.5g}"
                values.append(v)
            self.batch_tree.insert("", "end", values=values)

    def on_batch_tree_double_click(self, event):
        if self.batch_tree.identify_region(event.x, event.y) != "heading":
            return
        column_id = self.batch_tree.identify_column(event.x)
        if not column_id:
            return
        try:
            column_index = int(column_id.replace("#", "")) - 1
            column_name = self.batch_tree["columns"][column_index]
        except Exception:
            return
        self.sort_preview_by_column(column_name)

    def sort_preview_by_column(self, column_name: str):
        if self.preview_df is None or column_name not in self.preview_df.columns:
            return
        ascending = not self.preview_sort_ascending if self.preview_sort_column == column_name else True
        sorted_df = self.preview_df.copy()
        sort_values = pd.to_numeric(sorted_df[column_name], errors="coerce")
        if sort_values.notna().any():
            sorted_df = sorted_df.assign(_sort_key=sort_values).sort_values(
                "_sort_key",
                ascending=ascending,
                na_position="last",
                kind="mergesort",
            ).drop(columns="_sort_key")
        else:
            sorted_df = sorted_df.sort_values(
                column_name,
                ascending=ascending,
                na_position="last",
                kind="mergesort",
                key=lambda series: series.astype(str).str.lower(),
            )
        self.preview_sort_column = column_name
        self.preview_sort_ascending = ascending
        self.preview_df = sorted_df
        self.render_preview_table(sorted_df)

    def generate_batch_graphs(self, show_errors: bool = True):
        try:
            if self.last_batch_result is None:
                raise ValueError("Run batch prediction first.")
            out_dir = os.path.join(self.validated_output_dir(), "graphs")
            self.set_status("Generating graphs...")
            paths = save_all_batch_graphs(self.last_batch_result, out_dir, self.ad_calculator, self.last_batch_fp)
            self.ui(self.refresh_graph_list, paths)
            if paths:
                self.set_status(f"Graphs generated: {len(paths)}")
            return paths
        except Exception as e:
            if show_errors:
                self.show_error("Graph generation failed", e)
            return []

    def generate_graphs_clicked(self):
        def job():
            paths = self.generate_batch_graphs(show_errors=True)
            if paths:
                self.set_status(f"Graphs saved: {os.path.join(self.validated_output_dir(), 'graphs')}")
        self.run_threaded(job)

    def refresh_graph_list(self, paths=None):
        graph_dir = os.path.join(self.output_dir_var.get(), "graphs")
        if paths is None:
            if os.path.exists(graph_dir):
                paths = [os.path.join(graph_dir, f) for f in os.listdir(graph_dir) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
            else:
                paths = []
        paths = [path for path in paths if os.path.basename(path).lower() != "single_ad_pca_plot.png"]
        paths = sorted(paths, key=lambda path: (0 if os.path.basename(path).lower() == "ad_pca_plot.png" else 1, os.path.basename(path).lower()))
        self.graph_paths_by_name = {os.path.basename(path): path for path in paths}
        names = list(self.graph_paths_by_name.keys())
        self.graph_combo["values"] = names
        if paths:
            selected = "ad_pca_plot.png" if "ad_pca_plot.png" in self.graph_paths_by_name else names[0]
            self.graph_combo.set(selected)
            self.display_selected_graph()
            self.update_idletasks()
            self.set_status(f"Graph preview loaded: {selected}")
        else:
            self.graph_combo.set("")
            self.graph_label.configure(text="AD graph will appear after batch prediction.", image="")

    def display_selected_graph(self):
        if not PIL_AVAILABLE:
            self.graph_label.configure(text="Pillow is not installed; graph preview unavailable.", image="")
            return
        path = self.graph_paths_by_name.get(self.graph_combo.get(), self.graph_combo.get())
        if not path or not os.path.exists(path):
            self.graph_label.configure(text="Selected graph file was not found.", image="")
            return
        try:
            photo = self.make_preview_photo(path, self.batch_graph_preview_size)
            self.loaded_images.append(photo)
            self.graph_label.configure(image=photo, text="")
        except Exception as e:
            self.graph_label.configure(text=f"Graph preview failed:\n{e}", image="")


class StructureEditorDialog(tk.Toplevel):
    ATOMS = ["C", "N", "O", "S", "F", "Cl", "Br", "I", "P"]
    BONDS = {"Single": 1, "Double": 2, "Triple": 3}

    def __init__(self, parent: tk.Tk, smiles_var: tk.StringVar):
        super().__init__(parent)
        self.parent = parent
        self.smiles_var = smiles_var
        self.title("Draw structure")
        self.geometry("680x520")
        self.minsize(620, 460)
        self.transient(parent)
        self.grab_set()

        self.atoms = []
        self.bonds = []
        self.selected_atom = None
        self.undo_stack = []
        self.atom_var = tk.StringVar(value="C")
        self.bond_var = tk.StringVar(value="Single")

        self._build_ui()
        self._draw()

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self)
        toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)

        ttk.Label(toolbar, text="Atom").pack(side="left", padx=(0, 4))
        ttk.Combobox(toolbar, textvariable=self.atom_var, values=self.ATOMS, state="readonly", width=6).pack(side="left", padx=(0, 10))

        ttk.Label(toolbar, text="Bond").pack(side="left", padx=(0, 4))
        ttk.Combobox(toolbar, textvariable=self.bond_var, values=list(self.BONDS.keys()), state="readonly", width=8).pack(side="left", padx=(0, 10))

        ttk.Button(toolbar, text="Undo", command=self.undo).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Clear", command=self.clear).pack(side="left", padx=3)
        ttk.Button(toolbar, text="Save SMILES", command=self.save_smiles).pack(side="right", padx=3)
        ttk.Button(toolbar, text="Cancel", command=self.destroy).pack(side="right", padx=3)

        self.canvas = tk.Canvas(self, bg="white", highlightthickness=1, highlightbackground="#d0d7de")
        self.canvas.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 6))
        self.canvas.bind("<Button-1>", self.on_click)

        help_text = "Click empty space to add an atom. Select an atom, then click another atom or empty space to create a bond."
        ttk.Label(self, text=help_text, anchor="w").grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 8))

    def push_undo(self):
        atoms = [dict(atom) for atom in self.atoms]
        bonds = [tuple(bond) for bond in self.bonds]
        self.undo_stack.append((atoms, bonds, self.selected_atom))
        if len(self.undo_stack) > 30:
            self.undo_stack.pop(0)

    def undo(self):
        if not self.undo_stack:
            return
        self.atoms, self.bonds, self.selected_atom = self.undo_stack.pop()
        self._draw()

    def clear(self):
        if not self.atoms and not self.bonds:
            return
        self.push_undo()
        self.atoms = []
        self.bonds = []
        self.selected_atom = None
        self._draw()

    def on_click(self, event):
        idx = self.find_atom(event.x, event.y)
        if idx is None:
            self.add_atom(event.x, event.y)
            return

        if self.selected_atom is None:
            self.selected_atom = idx
        elif self.selected_atom == idx:
            self.selected_atom = None
        else:
            self.push_undo()
            self.add_or_update_bond(self.selected_atom, idx)
            self.selected_atom = idx
        self._draw()

    def add_atom(self, x, y):
        self.push_undo()
        self.atoms.append({"x": float(x), "y": float(y), "symbol": self.atom_var.get()})
        new_idx = len(self.atoms) - 1
        if self.selected_atom is not None:
            self.add_or_update_bond(self.selected_atom, new_idx)
        self.selected_atom = new_idx
        self._draw()

    def add_or_update_bond(self, i, j):
        if i == j:
            return
        order = self.BONDS.get(self.bond_var.get(), 1)
        a, b = sorted((int(i), int(j)))
        for idx, (x, y, _) in enumerate(self.bonds):
            if x == a and y == b:
                self.bonds[idx] = (a, b, order)
                return
        self.bonds.append((a, b, order))

    def find_atom(self, x, y):
        for idx, atom in enumerate(self.atoms):
            if math.hypot(atom["x"] - x, atom["y"] - y) <= 18:
                return idx
        return None

    def _draw(self):
        self.canvas.delete("all")
        for i, j, order in self.bonds:
            self.draw_bond(self.atoms[i], self.atoms[j], order)
        for idx, atom in enumerate(self.atoms):
            x, y = atom["x"], atom["y"]
            if idx == self.selected_atom:
                self.canvas.create_oval(x - 18, y - 18, x + 18, y + 18, outline="#0969da", width=2)
            self.canvas.create_oval(x - 12, y - 12, x + 12, y + 12, fill="white", outline="#57606a")
            self.canvas.create_text(x, y, text=atom["symbol"], fill="#24292f", font=("Segoe UI", 10, "bold"))

    def draw_bond(self, atom_a, atom_b, order):
        x1, y1 = atom_a["x"], atom_a["y"]
        x2, y2 = atom_b["x"], atom_b["y"]
        dx, dy = x2 - x1, y2 - y1
        length = max(math.hypot(dx, dy), 1.0)
        ox, oy = -dy / length * 4.0, dx / length * 4.0

        if order == 1:
            offsets = [(0, 0)]
        elif order == 2:
            offsets = [(-ox, -oy), (ox, oy)]
        else:
            offsets = [(-ox * 1.4, -oy * 1.4), (0, 0), (ox * 1.4, oy * 1.4)]

        for off_x, off_y in offsets:
            self.canvas.create_line(x1 + off_x, y1 + off_y, x2 + off_x, y2 + off_y, fill="#24292f", width=2)

    def save_smiles(self):
        if not self.atoms:
            messagebox.showwarning("No structure", "Draw at least one atom before saving.", parent=self)
            return
        try:
            smiles = self.to_smiles()
            self.smiles_var.set(smiles)
            self.parent.set_status(f"Structure drawing saved as SMILES: {smiles}")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Invalid structure", str(e), parent=self)

    def to_smiles(self):
        from rdkit import Chem

        rw_mol = Chem.RWMol()
        for atom in self.atoms:
            rw_mol.AddAtom(Chem.Atom(atom["symbol"]))

        bond_types = {
            1: Chem.BondType.SINGLE,
            2: Chem.BondType.DOUBLE,
            3: Chem.BondType.TRIPLE,
        }
        for i, j, order in self.bonds:
            if rw_mol.GetBondBetweenAtoms(int(i), int(j)) is None:
                rw_mol.AddBond(int(i), int(j), bond_types.get(int(order), Chem.BondType.SINGLE))

        mol = rw_mol.GetMol()
        Chem.SanitizeMol(mol)
        return Chem.MolToSmiles(mol)
