import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue
import time
from typing import Optional

# Pastikan path App berada di sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from main import (
    run_automation,
    list_date_folders,
    resolve_work_dir,
    PROJECT_ROOT,
    PipelineMode,
    MODE_DESCRIPTIONS,
    MODE_DEFAULT_STAGES,
    StageStatus,
    StageMetric,
    DailyDataError,
    FileLockError,
    StageExecutionError
)
from item_inspector import inspect_item, ItemInspectionResult, SOOrderItem
from recheck_sp import run_recheck_sp, RecheckSPResult, RecheckItem
import zlib
import struct
import base64


def _create_png_b64(width: int, height: int, rgba_pixels: list) -> str:
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack('>I', len(data)) + tag + data + struct.pack('>I', zlib.crc32(tag + data))
    raw = bytearray()
    for y in range(height):
        raw.append(0)
        for x in range(width):
            raw.extend(rgba_pixels[y * width + x])
    ihdr = struct.pack('>IIBBBBB', width, height, 8, 6, 0, 0, 0)
    png = b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', ihdr) + chunk(b'IDAT', zlib.compress(bytes(raw))) + chunk(b'IEND', b'')
    return base64.b64encode(png).decode('ascii')


class IconLibrary:
    """
    Koleksi ikon vektor crisp 16x16 PNG tanpa library luar.
    Menggantikan penggunaan emoticon/emoji dengan ikon profesional standard enterprise.
    """
    def __init__(self):
        self._cache = {}

    def get_play_icon(self) -> tk.PhotoImage:
        if "play" not in self._cache:
            w, h = 16, 16
            px = [[0, 0, 0, 0] for _ in range(w * h)]
            for y in range(3, 13):
                max_x = int(4 + (y - 3) * 0.9) if y <= 7 else int(4 + (12 - y) * 0.9)
                for x in range(4, min(14, max_x + 1)):
                    px[y * w + x] = [255, 255, 255, 255]
            self._cache["play"] = tk.PhotoImage(data=_create_png_b64(w, h, px))
        return self._cache["play"]

    def get_search_icon(self, color=(255, 255, 255)) -> tk.PhotoImage:
        key = f"search_{color}"
        if key not in self._cache:
            w, h = 16, 16
            px = [[0, 0, 0, 0] for _ in range(w * h)]
            r, g, b = color
            for y in range(2, 9):
                for x in range(2, 9):
                    dx, dy = x - 5.5, y - 5.5
                    d_sq = dx*dx + dy*dy
                    if 4.5 <= d_sq <= 10.5:
                        px[y * w + x] = [r, g, b, 255]
            for i in range(5):
                px[(9 + i) * w + (9 + i)] = [r, g, b, 255]
                if 9 + i + 1 < 16:
                    px[(9 + i) * w + (9 + i + 1)] = [r, g, b, 255]
            self._cache[key] = tk.PhotoImage(data=_create_png_b64(w, h, px))
        return self._cache[key]

    def get_box_icon(self) -> tk.PhotoImage:
        if "box" not in self._cache:
            w, h = 16, 16
            px = [[0, 0, 0, 0] for _ in range(w * h)]
            for y in range(3, 13):
                for x in range(2, 14):
                    if y in (3, 7, 12) or x in (2, 13) or (y > 3 and x == 7):
                        px[y * w + x] = [255, 255, 255, 255]
            self._cache["box"] = tk.PhotoImage(data=_create_png_b64(w, h, px))
        return self._cache["box"]

    def get_copy_icon(self) -> tk.PhotoImage:
        if "copy" not in self._cache:
            w, h = 16, 16
            px = [[0, 0, 0, 0] for _ in range(w * h)]
            for y in range(2, 11):
                for x in range(2, 10):
                    if y in (2, 10) or x in (2, 9):
                        px[y * w + x] = [100, 116, 139, 255]
            for y in range(5, 14):
                for x in range(5, 14):
                    if y in (5, 13) or x in (5, 13):
                        px[y * w + x] = [15, 23, 42, 255]
            self._cache["copy"] = tk.PhotoImage(data=_create_png_b64(w, h, px))
        return self._cache["copy"]

    def get_refresh_icon(self) -> tk.PhotoImage:
        if "refresh" not in self._cache:
            w, h = 16, 16
            px = [[0, 0, 0, 0] for _ in range(w * h)]
            for y in range(2, 14):
                for x in range(2, 14):
                    dx, dy = x - 7.5, y - 7.5
                    d_sq = dx*dx + dy*dy
                    if 10.0 <= d_sq <= 22.0:
                        if not (x > 7 and y > 7 and x < 12):
                            px[y * w + x] = [71, 85, 105, 255]
            for (ax, ay) in [(11, 4), (12, 4), (13, 4), (13, 5), (13, 6)]:
                px[ay * w + ax] = [71, 85, 105, 255]
            self._cache["refresh"] = tk.PhotoImage(data=_create_png_b64(w, h, px))
        return self._cache["refresh"]

    def get_check_icon(self) -> tk.PhotoImage:
        if "check" not in self._cache:
            w, h = 16, 16
            px = [[0, 0, 0, 0] for _ in range(w * h)]
            pts = [(2, 8), (3, 9), (4, 10), (5, 11), (6, 12), (7, 11), (8, 10), (9, 9), (10, 8), (11, 7), (12, 6), (13, 5)]
            for (cx, cy) in pts:
                for dy in [0, 1]:
                    if 0 <= cy + dy < h:
                        px[(cy + dy) * w + cx] = [255, 255, 255, 255]
            self._cache["check"] = tk.PhotoImage(data=_create_png_b64(w, h, px))
        return self._cache["check"]


class DailyDataGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Daily Data - Sistem Otomasi Pipeline")
        self.root.geometry("1020x740")
        self.root.minsize(850, 600)

        # Styling tema
        style = ttk.Style()
        style.theme_use('clam')

        self.icons = IconLibrary()
        self.log_queue = queue.Queue()
        self.stage_queue = queue.Queue()
        self.is_running = False
        self.stage_widgets = {}

        # Definisi tahapan pipeline
        self.stage_defs = [
            (1, "01 Init"),
            (2, "02 Stok"),
            (3, "03 Trn"),
            (4, "04 SOOUT"),
            (5, "05 POOUT"),
            (6, "06 Format"),
            (7, "07 Bukmas"),
            (8, "08 DO Sum"),
            (9, "09 Recheck")
        ]

        self._build_ui()
        self._check_queues()

    def _build_ui(self):
        # 1. HEADER (Desain Minimalis Elegan ala Google/Apple)
        header_frame = tk.Frame(self.root, bg="#0f172a", padx=20, pady=14)
        header_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            header_frame,
            text="DAILY DATA",
            font=("Helvetica", 16, "bold"),
            fg="#ffffff",
            bg="#0f172a"
        )
        title_lbl.pack(anchor="w")

        subtitle_lbl = tk.Label(
            header_frame,
            text="Enterprise Pipeline • Stok, FIFO, Buku Masuk, Format Cabang & DO Summary",
            font=("Helvetica", 9),
            fg="#94a3b8",
            bg="#0f172a"
        )
        subtitle_lbl.pack(anchor="w")

        # 2. CONFIG / TARGET FOLDER SELECTION
        config_frame = ttk.LabelFrame(self.root, text=" Target Folder Harian ", padding=10)
        config_frame.pack(fill=tk.X, padx=15, pady=(10, 5))

        folders = [os.path.basename(d) for d in list_date_folders()]
        self.dir_var = tk.StringVar(value=folders[0] if folders else "")

        self.dir_combo = ttk.Combobox(config_frame, textvariable=self.dir_var, values=folders, font=("Helvetica", 10))
        self.dir_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        refresh_btn = tk.Button(
            config_frame,
            text=" Refresh",
            image=self.icons.get_refresh_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 9),
            padx=8,
            pady=3,
            command=self._refresh_folders
        )
        refresh_btn.pack(side=tk.RIGHT, padx=(5, 0))

        browse_btn = ttk.Button(config_frame, text="Pilih Folder...", command=self._browse_dir)
        browse_btn.pack(side=tk.RIGHT)

        # 2.5 WORKFLOW / TAHAPAN SELECTION (LONCAT TAHAP / PRESET)
        workflow_frame = ttk.LabelFrame(self.root, text=" Alur Kerja & Pilihan Tahap (Workflow) ", padding=8)
        workflow_frame.pack(fill=tk.X, padx=15, pady=(2, 6))

        self.modes_list = [
            (PipelineMode.ALL, "[1] Semua Tahap (1 - 9) [Full Pipeline]"),
            (PipelineMode.POSISI_STOK, "[2] Posisi Stok + Recheck SP [Mode Pagi - Validasi Posting Kemarin]"),
            (PipelineMode.RECHECK_SP, "[3] Recheck SP Kemarin Saja [Validasi Posting & Cek Selisih Fisik]"),
            (PipelineMode.BUKU_MASUK, "[4] Buku Masuk & FIFO (Tahap 7 Saja) [Mode Siang - Ekspor Data Baru]"),
            (PipelineMode.DO_SUMMARY, "[5] DO Summary External (Tahap 8 Saja) [Mode Sore - Rekapitulasi CSV]"),
            (PipelineMode.CUSTOM, "[6] Kustom (Pilih Tahapan Manual)")
        ]

        wf_top_row = tk.Frame(workflow_frame)
        wf_top_row.pack(fill=tk.X)

        tk.Label(wf_top_row, text="Mode Eksekusi:", font=("Helvetica", 9, "bold")).pack(side=tk.LEFT, padx=(0, 8))

        self.mode_var = tk.StringVar(value=self.modes_list[0][1])
        self.mode_combo = ttk.Combobox(
            wf_top_row,
            textvariable=self.mode_var,
            values=[m[1] for m in self.modes_list],
            state="readonly",
            font=("Helvetica", 9)
        )
        self.mode_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.mode_combo.bind("<<ComboboxSelected>>", self._on_mode_selected)

        # Container Checkbox Custom (tampil hanya saat mode Kustom aktif)
        self.custom_stages_frame = tk.Frame(workflow_frame)
        self.stage_check_vars = {}
        for idx, short_name in self.stage_defs:
            v = tk.BooleanVar(value=True)
            chk = ttk.Checkbutton(
                self.custom_stages_frame,
                text=short_name,
                variable=v,
                command=self._on_custom_stages_toggled
            )
            chk.pack(side=tk.LEFT, padx=3)
            self.stage_check_vars[idx] = v

        # 3. ACTION CONTROLS (RUN, UPDATE BUKU MASUK, CEK ITEM)
        action_frame = tk.Frame(self.root, padx=15, pady=6)
        action_frame.pack(fill=tk.X)

        # Tombol RUN bergaya modern minimalis
        self.run_btn = tk.Button(
            action_frame,
            text=" RUN",
            image=self.icons.get_play_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 11, "bold"),
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            padx=20,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._start_pipeline
        )
        self.run_btn.pack(side=tk.LEFT)

        # Tombol Update Buku Masuk Cepat
        self.bukmas_btn = tk.Button(
            action_frame,
            text=" Update Buku Masuk",
            image=self.icons.get_box_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 10, "bold"),
            bg="#0f766e",
            fg="white",
            activebackground="#115e59",
            activeforeground="white",
            padx=14,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._quick_update_bukmas
        )
        self.bukmas_btn.pack(side=tk.LEFT, padx=(8, 0))

        # Tombol Cek Item & Antrean SO FIFO
        self.inspect_btn = tk.Button(
            action_frame,
            text=" Cek Item & Antrean",
            image=self.icons.get_search_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 10, "bold"),
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            padx=14,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._open_item_inspector
        )
        self.inspect_btn.pack(side=tk.LEFT, padx=(8, 0))

        # Tombol Recheck SP Kemarin
        self.recheck_btn = tk.Button(
            action_frame,
            text=" Recheck SP Kemarin",
            image=self.icons.get_check_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 10, "bold"),
            bg="#4f46e5",
            fg="white",
            activebackground="#4338ca",
            activeforeground="white",
            padx=14,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._open_recheck_sp
        )
        self.recheck_btn.pack(side=tk.LEFT, padx=(8, 0))

        # Tombol Proses DO Summary Sore
        self.dosum_btn = tk.Button(
            action_frame,
            text=" DO Summary Sore",
            image=self.icons.get_box_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 10, "bold"),
            bg="#d97706",
            fg="white",
            activebackground="#b45309",
            activeforeground="white",
            padx=14,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._quick_process_do_summary
        )
        self.dosum_btn.pack(side=tk.LEFT, padx=(8, 0))

        self.progress_bar = ttk.Progressbar(action_frame, mode='indeterminate')
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(15, 0))

        # 4. TIMELINE TRACKER (Visual Timeline seperti Video Editor)
        timeline_frame = ttk.LabelFrame(self.root, text=" Timeline Pipeline Eksekusi ", padding=8)
        timeline_frame.pack(fill=tk.X, padx=15, pady=(4, 8))

        timeline_badges_frame = tk.Frame(timeline_frame)
        timeline_badges_frame.pack(fill=tk.X)

        for idx, short_name in self.stage_defs:
            badge_lbl = tk.Label(
                timeline_badges_frame,
                text=f"[WAIT] {short_name}",
                font=("Consolas", 8, "bold"),
                bg="#e2e8f0",
                fg="#475569",
                padx=5,
                pady=4,
                relief=tk.GROOVE
            )
            badge_lbl.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            self.stage_widgets[idx] = badge_lbl

        self._update_badges_preview()

        # 5. CONSOLE LOG REAL-TIME (Syntax Color Coded)
        log_frame = ttk.LabelFrame(self.root, text=" Log Console ", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 10))

        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="#020617",
            fg="#f8fafc",
            insertbackground="white",
            padx=8,
            pady=8
        )
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Tag warna syntax log
        self.log_text.tag_config("STAGE", foreground="#38bdf8", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("SUCCESS", foreground="#4ade80", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("ERROR", foreground="#f87171", font=("Consolas", 9, "bold"))
        self.log_text.tag_config("WARN", foreground="#facc15")
        self.log_text.tag_config("INFO", foreground="#e2e8f0")

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 6. STATUS BAR
        self.status_var = tk.StringVar(value="Siap dijalankan.")
        status_bar = tk.Label(
            self.root,
            textvariable=self.status_var,
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Helvetica", 8),
            padx=10,
            pady=3,
            bg="#f8fafc",
            fg="#475569"
        )
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _browse_dir(self):
        selected = filedialog.askdirectory(initialdir=PROJECT_ROOT)
        if selected:
            self.dir_var.set(selected)

    def _refresh_folders(self):
        folders = [os.path.basename(d) for d in list_date_folders()]
        self.dir_combo.configure(values=folders)
        if folders and not self.dir_var.get():
            self.dir_var.set(folders[0])

    def _append_log(self, text: str):
        self.log_queue.put(text)

    def _stage_update_callback(self, stage_metric: StageMetric):
        self.stage_queue.put(stage_metric)

    def _check_queues(self):
        # 1. Cek antrean log
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            tag = "INFO"
            if "[STAGE]" in msg:
                tag = "STAGE"
            elif "[SUCCESS]" in msg:
                tag = "SUCCESS"
            elif "[ERROR]" in msg:
                tag = "ERROR"
            elif "[WARN]" in msg or "[PERINGATAN]" in msg:
                tag = "WARN"

            self.log_text.insert(tk.END, msg + "\n", tag)
            self.log_text.see(tk.END)

        # 2. Cek antrean timeline visual
        while not self.stage_queue.empty():
            stg: StageMetric = self.stage_queue.get_nowait()
            widget = self.stage_widgets.get(stg.index)
            if widget:
                short_name = self.stage_defs[stg.index - 1][1]
                if stg.status == StageStatus.RUNNING:
                    widget.configure(text=f"[RUN] {short_name}", bg="#0284c7", fg="#ffffff")
                elif stg.status == StageStatus.SUCCESS:
                    widget.configure(text=f"[OK] {short_name}", bg="#15803d", fg="#ffffff")
                elif stg.status == StageStatus.SKIPPED:
                    widget.configure(text=f"[SKIP] {short_name}", bg="#94a3b8", fg="#ffffff")
                elif stg.status == StageStatus.FAILED:
                    widget.configure(text=f"[ERR] {short_name}", bg="#dc2626", fg="#ffffff")
                else:
                    widget.configure(text=f"[WAIT] {short_name}", bg="#e2e8f0", fg="#475569")

        self.root.after(80, self._check_queues)

    def _get_current_mode_key(self) -> PipelineMode:
        selected_text = self.mode_var.get()
        for k, text in self.modes_list:
            if text == selected_text:
                return k
        return PipelineMode.ALL

    def _get_active_stages(self) -> list:
        mode_key = self._get_current_mode_key()
        if mode_key == PipelineMode.CUSTOM:
            active = [idx for idx, var in self.stage_check_vars.items() if var.get()]
            if 1 not in active and active:
                active.insert(0, 1)
            return active
        return MODE_DEFAULT_STAGES.get(mode_key, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10])

    def _on_mode_selected(self, event=None):
        mode_key = self._get_current_mode_key()
        if mode_key == PipelineMode.CUSTOM:
            self.custom_stages_frame.pack(fill=tk.X, pady=(6, 2))
        else:
            self.custom_stages_frame.pack_forget()

        self._update_badges_preview()

    def _on_custom_stages_toggled(self):
        self._update_badges_preview()

    def _update_badges_preview(self):
        active_stages = self._get_active_stages()
        for idx, short_name in self.stage_defs:
            w = self.stage_widgets.get(idx)
            if w:
                if idx in active_stages:
                    w.configure(text=f"[WAIT] {short_name}", bg="#e2e8f0", fg="#0f172a")
                else:
                    w.configure(text=f"[SKIP] {short_name}", bg="#f1f5f9", fg="#94a3b8")

    def _reset_timeline(self):
        self._update_badges_preview()

    def _start_pipeline(self):
        if self.is_running:
            return

        target_str = self.dir_var.get().strip()
        if not target_str:
            messagebox.showerror("Error", "Pilih atau ketik nama folder harian terlebih dahulu.")
            return

        try:
            target_dir = resolve_work_dir(target_str)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        mode_key = self._get_current_mode_key()
        active_stages = self._get_active_stages()
        if not active_stages:
            messagebox.showwarning("Peringatan", "Pilih minimal satu tahapan untuk dijalankan.")
            return

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.bukmas_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.recheck_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.dosum_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.progress_bar.start(10)
        self.status_var.set("Menjalankan pipeline Daily Data... Harap tunggu.")
        self.log_text.delete(1.0, tk.END)
        self._reset_timeline()

        threading.Thread(target=self._run_thread, args=(target_dir, mode_key, active_stages), daemon=True).start()

    def _quick_update_bukmas(self):
        """
        Eksekusi 1-klik untuk menerima dan mengekspor buku masuk baru secara inkremental (Tahap 1 & 7).
        """
        if self.is_running:
            return

        target_str = self.dir_var.get().strip()
        if not target_str:
            messagebox.showerror("Error", "Pilih atau ketik nama folder harian terlebih dahulu.")
            return

        try:
            target_dir = resolve_work_dir(target_str)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.bukmas_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.recheck_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.dosum_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.progress_bar.start(10)
        self.status_var.set("Memeriksa dan memperbarui data Buku Masuk baru...")
        self.log_text.delete(1.0, tk.END)
        self._reset_timeline()

        threading.Thread(target=self._run_thread, args=(target_dir, PipelineMode.BUKU_MASUK, [1, 7]), daemon=True).start()

    def _quick_process_do_summary(self):
        """
        Eksekusi 1-klik untuk DO Summary External (Tahap 1, 8 & 10).
        Secara otomatis mendeteksi file DO Summary hari ini di Downloads / Folder Kerja.
        """
        if self.is_running:
            return

        target_str = self.dir_var.get().strip()
        if not target_str:
            messagebox.showerror("Error", "Pilih atau ketik nama folder harian terlebih dahulu.")
            return

        try:
            target_dir = resolve_work_dir(target_str)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.bukmas_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.recheck_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.dosum_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.progress_bar.start(10)
        self.status_var.set("Memproses DO Summary External... Harap tunggu.")
        self.log_text.delete(1.0, tk.END)
        self._reset_timeline()

        threading.Thread(target=self._run_thread, args=(target_dir, PipelineMode.DO_SUMMARY, [1, 8]), daemon=True).start()

    def _run_thread(self, target_dir: str, mode_key: PipelineMode, active_stages: list):
        try:
            timeline = run_automation(
                work_dir=target_dir,
                mode=mode_key,
                selected_stages=active_stages,
                progress_callback=self._append_log,
                stage_callback=self._stage_update_callback
            )
            self._append_log("\n[SUKSES] Seluruh tahapan pipeline yang dipilih berhasil diselesaikan 100%.")
            self.root.after(0, lambda: messagebox.showinfo(
                "Sukses",
                "Proses Daily Data berhasil diselesaikan 100% tanpa kendala."
            ))
            self.status_var.set("Selesai dengan sukses.")
        except Exception as e:
            self.status_var.set("Pipeline gagal. Periksa diagnosis error.")
            self.root.after(0, lambda exc=e: self._show_error_dialog(exc))
        finally:
            self.is_running = False
            self.root.after(0, self._reset_ui)

    def _show_error_dialog(self, exc: Exception):
        """
        Dialog Error Diagnostik Enterprise:
        Menampilkan tahap kegagalan, solusi praktis, dan traceback teknis.
        """
        err_win = tk.Toplevel(self.root)
        err_win.title("Diagnosis Kesalahan Pipeline")
        err_win.geometry("640x440")
        err_win.minsize(550, 360)
        err_win.transient(self.root)
        err_win.grab_set()

        header = tk.Frame(err_win, bg="#dc2626", padx=15, pady=10)
        header.pack(fill=tk.X)

        stage_title = "Kesalahan Pipeline"
        if isinstance(exc, StageExecutionError):
            stage_title = f"Gagal pada Tahap [{exc.stage_idx}]: {exc.stage_name}"

        tk.Label(
            header,
            text="KESALAHAN EKSEKUSI PIPELINE",
            font=("Helvetica", 11, "bold"),
            fg="white",
            bg="#dc2626"
        ).pack(anchor="w")

        tk.Label(
            header,
            text=stage_title,
            font=("Helvetica", 9),
            fg="#fee2e2",
            bg="#dc2626"
        ).pack(anchor="w")

        body = tk.Frame(err_win, padx=15, pady=10)
        body.pack(fill=tk.BOTH, expand=True)

        # Penjelasan Error
        tk.Label(body, text="Penyebab:", font=("Helvetica", 9, "bold")).pack(anchor="w")
        msg_lbl = tk.Label(body, text=str(exc), font=("Helvetica", 9), fg="#991b1b", wraplength=580, justify=tk.LEFT)
        msg_lbl.pack(anchor="w", pady=(2, 6))

        # Solusi / Tips
        tip = getattr(exc, "resolution_tip", "")
        if tip:
            tk.Label(body, text="Solusi / Rekomendasi:", font=("Helvetica", 9, "bold")).pack(anchor="w")
            tip_lbl = tk.Label(body, text=tip, font=("Helvetica", 9), fg="#166534", wraplength=580, justify=tk.LEFT)
            tip_lbl.pack(anchor="w", pady=(2, 6))

        # Rincian Teknis
        tk.Label(body, text="Detail Log / Traceback:", font=("Helvetica", 8, "bold")).pack(anchor="w", pady=(4, 0))
        tb_text = tk.Text(body, height=8, font=("Consolas", 8), bg="#f8fafc", fg="#334155")
        tb_text.pack(fill=tk.BOTH, expand=True, pady=4)

        tb_str = getattr(exc, "traceback_str", "") or str(exc)
        tb_text.insert(tk.END, tb_str)
        tb_text.configure(state=tk.DISABLED)

        # Tombol
        btn_frame = tk.Frame(err_win, padx=15, pady=10)
        btn_frame.pack(fill=tk.X)

        def copy_tb():
            self.root.clipboard_clear()
            self.root.clipboard_append(f"{str(exc)}\n\n{tb_str}")
            messagebox.showinfo("Tersalin", "Detail error berhasil disalin ke clipboard.")

        copy_btn = ttk.Button(btn_frame, text="Salin Rincian Error", command=copy_tb)
        copy_btn.pack(side=tk.LEFT)

        close_btn = ttk.Button(btn_frame, text="Tutup", command=err_win.destroy)
        close_btn.pack(side=tk.RIGHT)

    def _reset_ui(self):
        self.run_btn.configure(state=tk.NORMAL, bg="#16a34a")
        self.bukmas_btn.configure(state=tk.NORMAL, bg="#0f766e")
        self.recheck_btn.configure(state=tk.NORMAL, bg="#4f46e5")
        self.dosum_btn.configure(state=tk.NORMAL, bg="#d97706")
        self.progress_bar.stop()

    def _open_item_inspector(self):
        target_str = self.dir_var.get().strip()
        try:
            target_dir = resolve_work_dir(target_str) if target_str else PROJECT_ROOT
        except Exception:
            target_dir = PROJECT_ROOT
        ItemInspectorWindow(self.root, target_dir)

    def _open_recheck_sp(self):
        target_str = self.dir_var.get().strip()
        try:
            target_dir = resolve_work_dir(target_str) if target_str else PROJECT_ROOT
        except Exception:
            target_dir = PROJECT_ROOT
        RecheckSPWindow(self.root, target_dir)


class ItemInspectorWindow:
    """
    Jendela Inspeksi Item 360 Derajat & Simulasi Antrean FIFO SO:
    - Verifikasi real-time apakah item ada di Buku Masuk hari ini.
    - Menampilkan multi-stok (56 Makassar, In Transit TRN, 50 Pusat, Nasional, On Order).
    - Memetakan antrean SO 56-SOE terurut FIFO beserta hierarki sales (GL, ASM, MGR).
    - Melakukan simulasi pemenuhan stok (TERPENUHI, SEBAGIAN, MENUNGGU STOK).
    - Kemampuan menyalin data per item / per kolom / per sel secara individual.
    """
    def __init__(self, parent: tk.Tk, initial_dir: str):
        self.parent = parent
        self.work_dir = initial_dir
        self.current_result: Optional[ItemInspectionResult] = None
        self.is_searching = False
        self.search_queue = queue.Queue()
        self.icons = IconLibrary()
        self._current_badge_vals = {}

        self.win = tk.Toplevel(parent)
        self.win.title("Daily Data - Cek Item & Antrean SO (FIFO)")
        self.win.geometry("1020x760")
        self.win.minsize(900, 620)
        self.win.transient(parent)

        self._build_ui()
        self._check_search_queue()

    def _build_ui(self):
        # 1. HEADER
        header_frame = tk.Frame(self.win, bg="#0f172a", padx=18, pady=12)
        header_frame.pack(fill=tk.X)

        tk.Label(
            header_frame,
            text="PENGECEKAN ITEM & SIMULASI FIFO SO",
            font=("Helvetica", 14, "bold"),
            fg="#ffffff",
            bg="#0f172a"
        ).pack(anchor="w")

        tk.Label(
            header_frame,
            text="Verifikasi Buku Masuk Harian • Posisi Multi-Stok • Antrean SO FIFO & Tingkat Pemenuhan",
            font=("Helvetica", 9),
            fg="#94a3b8",
            bg="#0f172a"
        ).pack(anchor="w")

        # 2. SEARCH & CONTROLS FRAME
        ctrl_frame = ttk.LabelFrame(self.win, text=" Parameter Pencarian ", padding=10)
        ctrl_frame.pack(fill=tk.X, padx=15, pady=(8, 4))

        # Folder selector row
        row1 = tk.Frame(ctrl_frame)
        row1.pack(fill=tk.X, pady=(0, 6))

        tk.Label(row1, text="Folder Target:", font=("Helvetica", 9, "bold"), width=14, anchor="w").pack(side=tk.LEFT)
        folders = [os.path.basename(d) for d in list_date_folders()]
        curr_folder_name = os.path.basename(self.work_dir)
        self.dir_var = tk.StringVar(value=curr_folder_name if curr_folder_name in folders else (folders[0] if folders else ""))
        self.folder_combo = ttk.Combobox(row1, textvariable=self.dir_var, values=folders, width=24, state="readonly")
        self.folder_combo.pack(side=tk.LEFT, padx=(0, 10))

        # Search input row
        row2 = tk.Frame(ctrl_frame)
        row2.pack(fill=tk.X)

        tk.Label(row2, text="Item / Judul:", font=("Helvetica", 9, "bold"), width=14, anchor="w").pack(side=tk.LEFT)
        self.search_entry = ttk.Entry(row2, font=("Helvetica", 10))
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        self.search_entry.bind("<Return>", lambda e: self._start_search())
        self.search_entry.focus_set()

        self.btn_search = tk.Button(
            row2,
            text=" Cari & Analisis",
            image=self.icons.get_search_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 9, "bold"),
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            padx=12,
            pady=4,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._start_search
        )
        self.btn_search.pack(side=tk.LEFT, padx=(0, 5))

        btn_clear = ttk.Button(row2, text="Bersihkan", command=self._clear_search)
        btn_clear.pack(side=tk.LEFT)

        # Status / Helper Label
        self.search_status_lbl = tk.Label(
            ctrl_frame,
            text="Ketik kode item (misal: 0024200580) atau kata kunci judul buku, lalu klik Cari atau tekan Enter.",
            font=("Helvetica", 8),
            fg="#64748b",
            anchor="w"
        )
        self.search_status_lbl.pack(fill=tk.X, pady=(6, 0))

        # 3. RESULT CONTAINER
        self.res_frame = tk.Frame(self.win, padx=15, pady=4)
        self.res_frame.pack(fill=tk.BOTH, expand=True)

        # Placeholder Awal
        self.placeholder_lbl = tk.Label(
            self.res_frame,
            text="Silakan masukkan Kode Item atau Judul Buku di atas untuk memulai analisis.",
            font=("Helvetica", 11),
            fg="#94a3b8",
            pady=80
        )
        self.placeholder_lbl.pack(expand=True)

        # Container Data (sembunyi di awal)
        self.data_container = tk.Frame(self.res_frame)

        # Header Info Item & Status Buku Masuk
        self.info_frame = tk.Frame(self.data_container, bg="#f8fafc", relief=tk.GROOVE, bd=1, padx=12, pady=8)
        self.info_frame.pack(fill=tk.X, pady=(0, 6))

        top_info_row = tk.Frame(self.info_frame, bg="#f8fafc")
        top_info_row.pack(fill=tk.X)

        self.item_title_lbl = tk.Label(
            top_info_row,
            text="",
            font=("Helvetica", 12, "bold"),
            fg="#0f172a",
            bg="#f8fafc",
            anchor="w"
        )
        self.item_title_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Tombol Salin Data Individual Header
        self.header_copy_frame = tk.Frame(top_info_row, bg="#f8fafc")
        self.header_copy_frame.pack(side=tk.RIGHT)

        self.btn_cp_code = ttk.Button(
            self.header_copy_frame,
            text="Salin Kode",
            command=lambda: self._copy_single_value(self.current_result.item_code if self.current_result else "", "Kode Item")
        )
        self.btn_cp_code.pack(side=tk.LEFT, padx=2)

        self.btn_cp_title = ttk.Button(
            self.header_copy_frame,
            text="Salin Judul",
            command=lambda: self._copy_single_value(self.current_result.item_title if self.current_result else "", "Judul Buku")
        )
        self.btn_cp_title.pack(side=tk.LEFT, padx=2)

        # Row Status Buku Masuk
        self.bukmas_row = tk.Frame(self.info_frame, bg="#f8fafc")
        self.bukmas_row.pack(fill=tk.X, pady=(4, 0))

        self.bukmas_badge_lbl = tk.Label(
            self.bukmas_row,
            text="",
            font=("Helvetica", 9, "bold"),
            padx=8,
            pady=3,
            anchor="w"
        )
        self.bukmas_badge_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_cp_bm_qty = ttk.Button(
            self.bukmas_row,
            text="Salin Qty Masuk",
            command=lambda: self._copy_single_value(f"{self.current_result.buku_masuk_qty:,.0f}" if self.current_result else "0", "Qty Buku Masuk")
        )
        # Akan di-pack jika ada kuantiti buku masuk

        # Stock Badges Grid (6 Kartu Metrik Interaktif - Klik untuk Salin)
        badges_container = tk.Frame(self.data_container)
        badges_container.pack(fill=tk.X, pady=(0, 6))

        self.badge_labels = {}
        badge_specs = [
            ("mks", "Stok Fisik MKS (56)", "#eff6ff", "#1d4ed8"),
            ("trn", "In-Transit (TRN)", "#ecfeff", "#0e7490"),
            ("pst", "Stok Pusat (50)", "#eef2ff", "#4338ca"),
            ("nas", "Stok Nasional", "#f8fafc", "#334155"),
            ("so", "Total On Order SO", "#fff7ed", "#c2410c"),
            ("alloc", "Stok Siap Alokasi", "#f0fdf4", "#15803d"),
        ]

        for i, (key, label, bg_color, fg_color) in enumerate(badge_specs):
            card = tk.Frame(badges_container, bg=bg_color, relief=tk.GROOVE, bd=1, padx=8, pady=5, cursor="hand2")
            card.grid(row=0, column=i, sticky="nsew", padx=2)
            badges_container.grid_columnconfigure(i, weight=1)

            t_lbl = tk.Label(card, text=label, font=("Helvetica", 7, "bold"), bg=bg_color, fg="#64748b", cursor="hand2")
            t_lbl.pack(anchor="w")
            v_lbl = tk.Label(card, text="0 eks", font=("Helvetica", 10, "bold"), bg=bg_color, fg=fg_color, cursor="hand2")
            v_lbl.pack(anchor="w")
            hint_lbl = tk.Label(card, text="(klik salin)", font=("Helvetica", 6, "italic"), bg=bg_color, fg="#94a3b8", cursor="hand2")
            hint_lbl.pack(anchor="w")

            self.badge_labels[key] = v_lbl

            # Bind klik untuk salin nilai metrik
            def make_click_handler(k=key, l=label):
                return lambda e: self._on_badge_click(k, l)

            handler = make_click_handler()
            card.bind("<Button-1>", handler)
            t_lbl.bind("<Button-1>", handler)
            v_lbl.bind("<Button-1>", handler)
            hint_lbl.bind("<Button-1>", handler)

        # Fulfillment Status Banner
        self.fulfill_banner = tk.Label(
            self.data_container,
            text="",
            font=("Helvetica", 9, "bold"),
            padx=10,
            pady=5,
            relief=tk.GROOVE,
            anchor="w"
        )
        self.fulfill_banner.pack(fill=tk.X, pady=(0, 6))

        # Antrean SO Treeview Frame
        tree_frame = ttk.LabelFrame(self.data_container, text=" Antrean Sales Order (Urutan FIFO - Tanggal & Jam ASC) ", padding=6)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Toolbar Ringkas di atas tabel untuk salin cepat
        tree_toolbar = tk.Frame(tree_frame)
        tree_toolbar.pack(fill=tk.X, pady=(0, 4))

        tk.Label(
            tree_toolbar,
            text="Petunjuk: Dobel-klik baris untuk salin No SO. Klik kanan sel untuk opsi salin data per kolom.",
            font=("Helvetica", 8),
            fg="#64748b"
        ).pack(side=tk.LEFT)

        btn_cp_all_so = ttk.Button(
            tree_toolbar,
            text="Salin Semua SO (Delimit)",
            command=self._copy_all_sos_delimit
        )
        btn_cp_all_so.pack(side=tk.RIGHT, padx=2)

        btn_cp_row = ttk.Button(
            tree_toolbar,
            text="Salin Baris Terpilih",
            command=self._copy_selected_row
        )
        btn_cp_row.pack(side=tk.RIGHT, padx=2)

        cols = ("no", "so", "tgl", "sales", "gl", "asm", "mgr", "order", "alokasi", "status")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=8, selectmode="browse")

        col_configs = [
            ("no", "#", 35, "center"),
            ("so", "No SO", 110, "center"),
            ("tgl", "Waktu Order", 120, "center"),
            ("sales", "Nama Sales", 130, "w"),
            ("gl", "GL / REP", 100, "w"),
            ("asm", "ASM", 95, "w"),
            ("mgr", "MGR", 110, "w"),
            ("order", "Order", 60, "e"),
            ("alokasi", "Alokasi", 60, "e"),
            ("status", "Status Pemenuhan", 135, "center"),
        ]

        for cid, heading, width, align in col_configs:
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, anchor=align)

        tree_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Event Bindings pada Tabel
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Control-c>", self._on_tree_ctrl_c)

        # Treeview Tag Styling
        self.tree.tag_configure("TERPENUHI", background="#dcfce7", foreground="#166534")
        self.tree.tag_configure("SEBAGIAN", background="#fef9c3", foreground="#854d0e")
        self.tree.tag_configure("MENUNGGU_STOK", background="#fee2e2", foreground="#991b1b")

        # 4. BOTTOM ACTION BAR
        action_bar = tk.Frame(self.win, padx=15, pady=8)
        action_bar.pack(fill=tk.X)

        self.btn_copy = ttk.Button(
            action_bar,
            text=" Salin Ringkasan Lengkap",
            image=self.icons.get_copy_icon(),
            compound=tk.LEFT,
            state=tk.DISABLED,
            command=self._copy_to_clipboard
        )
        self.btn_copy.pack(side=tk.LEFT)

        self.copy_toast_lbl = tk.Label(action_bar, text="", font=("Helvetica", 9, "bold"), fg="#166534")
        self.copy_toast_lbl.pack(side=tk.LEFT, padx=12)

        btn_close = ttk.Button(action_bar, text="Tutup", command=self.win.destroy)
        btn_close.pack(side=tk.RIGHT)

    def _copy_single_value(self, val: str, label_name: str = ""):
        """
        Salin nilai tunggal ke clipboard dengan notifikasi visual instan.
        """
        if not val:
            return
        self.win.clipboard_clear()
        self.win.clipboard_append(val)
        display_text = f"{label_name}: {val}" if label_name else val
        if len(display_text) > 42:
            display_text = display_text[:39] + "..."
        self.copy_toast_lbl.configure(text=f"[TERSALIN] {display_text}")
        self.win.after(2800, lambda: self.copy_toast_lbl.configure(text=""))

    def _on_badge_click(self, key: str, label_name: str):
        val = str(self._current_badge_vals.get(key, "0"))
        self._copy_single_value(val, label_name)

    def _copy_selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if vals and vals[0] != "-":
            text = "\t".join(str(v) for v in vals)
            self._copy_single_value(text, f"Baris SO {vals[1]}")

    def _copy_all_sos_delimit(self):
        if not self.current_result or not self.current_result.so_orders:
            return
        sos = [o.sales_order for o in self.current_result.so_orders if o.sales_order and o.sales_order != "-"]
        if sos:
            delimit_str = ",".join(sos)
            self._copy_single_value(delimit_str, f"Semua {len(sos)} No SO (Delimit AX)")

    def _on_tree_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        vals = self.tree.item(row_id, "values")
        if not vals or vals[0] == "-":
            return

        col_idx = int(col_id.replace("#", "")) - 1 if col_id else 1
        clicked_cell = str(vals[col_idx]) if 0 <= col_idx < len(vals) else str(vals[1])
        col_names = ["#", "No SO", "Waktu Order", "Nama Sales", "GL / REP", "ASM", "MGR", "Order", "Alokasi", "Status"]
        clicked_col_name = col_names[col_idx] if 0 <= col_idx < len(col_names) else "Sel"

        menu = tk.Menu(self.win, tearoff=0)
        menu.add_command(
            label=f"Salin No SO: {vals[1]}",
            command=lambda: self._copy_single_value(vals[1], "No SO")
        )
        menu.add_command(
            label=f"Salin Nama Sales: {vals[3]}",
            command=lambda: self._copy_single_value(vals[3], "Sales")
        )
        menu.add_command(
            label=f"Salin Tanggal Order: {vals[2]}",
            command=lambda: self._copy_single_value(vals[2], "Tanggal")
        )
        menu.add_command(
            label=f"Salin Qty Order: {vals[7]} eks",
            command=lambda: self._copy_single_value(vals[7], "Qty Order")
        )
        menu.add_command(
            label=f"Salin {clicked_col_name}: {clicked_cell}",
            command=lambda: self._copy_single_value(clicked_cell, clicked_col_name)
        )
        menu.add_separator()
        menu.add_command(
            label="Salin Baris Terpilih (Format Excel Tab)",
            command=self._copy_selected_row
        )
        menu.add_command(
            label="Salin Semua No SO Antrean (Delimit Koma)",
            command=self._copy_all_sos_delimit
        )

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _on_tree_double_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if row_id:
            vals = self.tree.item(row_id, "values")
            if vals and vals[0] != "-":
                self._copy_single_value(vals[1], "No SO")

    def _on_tree_ctrl_c(self, event):
        self._copy_selected_row()

    def _clear_search(self):
        self.search_entry.delete(0, tk.END)
        self.data_container.pack_forget()
        self.placeholder_lbl.configure(
            text="Silakan masukkan Kode Item atau Judul Buku di atas untuk memulai analisis.",
            fg="#94a3b8"
        )
        self.placeholder_lbl.pack(expand=True)
        self.btn_copy.configure(state=tk.DISABLED)
        self.current_result = None
        self._current_badge_vals = {}
        self.search_status_lbl.configure(text="Form pencarian dibersihkan.", fg="#64748b")
        self.copy_toast_lbl.configure(text="")

    def _start_search(self):
        if self.is_searching:
            return
        query = self.search_entry.get().strip()
        if not query:
            self.search_status_lbl.configure(text="Harap masukkan kode item atau judul buku terlebih dahulu.", fg="#dc2626")
            return

        target_folder = self.dir_var.get().strip()
        try:
            target_dir = resolve_work_dir(target_folder)
        except Exception as e:
            self.search_status_lbl.configure(text=f"Folder target tidak valid: {e}", fg="#dc2626")
            return

        self.is_searching = True
        self.btn_search.configure(state=tk.DISABLED, bg="#94a3b8")
        self.search_status_lbl.configure(text=f"Sedang mencari dan menganalisis '{query}'...", fg="#0284c7")
        self.copy_toast_lbl.configure(text="")

        threading.Thread(target=self._search_worker, args=(query, target_dir), daemon=True).start()

    def _search_worker(self, query: str, target_dir: str):
        try:
            res = inspect_item(query, target_dir)
            self.search_queue.put((res, query, None))
        except Exception as e:
            self.search_queue.put((None, query, e))

    def _check_search_queue(self):
        try:
            if not self.win.winfo_exists():
                return
        except Exception:
            return

        while not self.search_queue.empty():
            res, query, exc = self.search_queue.get_nowait()
            self.is_searching = False
            self.btn_search.configure(state=tk.NORMAL, bg="#0284c7")
            if exc:
                self._on_search_error(exc)
            else:
                self._on_search_done(res, query)

        try:
            self.win.after(80, self._check_search_queue)
        except Exception:
            pass

    def _on_search_done(self, res: Optional[ItemInspectionResult], query: str):
        self.is_searching = False
        self.btn_search.configure(state=tk.NORMAL, bg="#0284c7")

        if not res:
            self.current_result = None
            self._current_badge_vals = {}
            self.data_container.pack_forget()
            self.placeholder_lbl.configure(
                text=f"Item '{query}' tidak ditemukan di Format Cabang, SOOUT, maupun Buku Masuk.",
                fg="#dc2626"
            )
            self.placeholder_lbl.pack(expand=True)
            self.search_status_lbl.configure(text=f"Pencarian '{query}' tidak membuahkan hasil.", fg="#dc2626")
            self.btn_copy.configure(state=tk.DISABLED)
            return

        self.current_result = res
        self._render_result(res)

    def _on_search_error(self, exc: Exception):
        self.is_searching = False
        self.btn_search.configure(state=tk.NORMAL, bg="#0284c7")
        self.search_status_lbl.configure(text=f"Terjadi kesalahan: {exc}", fg="#dc2626")
        messagebox.showerror("Error Pengecekan Item", str(exc), parent=self.win)

    def _render_result(self, res: ItemInspectionResult):
        self.placeholder_lbl.pack_forget()
        self.data_container.pack(fill=tk.BOTH, expand=True)

        # 1. Info Header
        self.item_title_lbl.configure(text=f"[{res.item_code}]  {res.item_title}")

        # Buku Masuk Badge
        if res.in_buku_masuk:
            files_str = ", ".join(res.buku_masuk_files)
            self.bukmas_badge_lbl.configure(
                text=f"[STATUS] DITEMUKAN DI BUKU MASUK: +{res.buku_masuk_qty:,.0f} eks  (File: {files_str})",
                bg="#dcfce7",
                fg="#15803d"
            )
            self.btn_cp_bm_qty.pack(side=tk.RIGHT, padx=4)
        else:
            self.bukmas_badge_lbl.configure(
                text="[STATUS] BUKU MASUK HARI INI: Belum Diterima / Tidak Ada",
                bg="#f1f5f9",
                fg="#64748b"
            )
            self.btn_cp_bm_qty.pack_forget()

        # 2. Stock Badges
        self._current_badge_vals = {
            "mks": int(res.stok_mks),
            "trn": int(res.stok_trn),
            "pst": int(res.stok_pst),
            "nas": int(res.stok_nas),
            "so": int(res.total_so_needed),
            "alloc": int(res.total_allocatable)
        }
        self.badge_labels["mks"].configure(text=f"{res.stok_mks:,.0f} eks")
        self.badge_labels["trn"].configure(text=f"{res.stok_trn:,.0f} eks")
        self.badge_labels["pst"].configure(text=f"{res.stok_pst:,.0f} eks")
        self.badge_labels["nas"].configure(text=f"{res.stok_nas:,.0f} eks")
        self.badge_labels["so"].configure(text=f"{res.total_so_needed:,.0f} eks")
        self.badge_labels["alloc"].configure(text=f"{res.total_allocatable:,.0f} eks")

        # 3. Fulfillment Banner
        if "100%" in res.fulfillment_summary or "TERPENUHI" in res.fulfillment_summary:
            banner_bg = "#dcfce7"
            banner_fg = "#166534"
        elif "SEBAGIAN" in res.fulfillment_summary:
            banner_bg = "#fef9c3"
            banner_fg = "#854d0e"
        elif "MENUNGGU" in res.fulfillment_summary:
            banner_bg = "#fee2e2"
            banner_fg = "#991b1b"
        else:
            banner_bg = "#f1f5f9"
            banner_fg = "#334155"

        self.fulfill_banner.configure(
            text=f"STATUS PEMENUHAN: {res.fulfillment_summary}",
            bg=banner_bg,
            fg=banner_fg
        )

        # 4. Populate Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not res.so_orders:
            self.tree.insert("", tk.END, values=("-", "-", "-", "Tidak ada antrean SO (On Order = 0)", "-", "-", "-", 0, 0, "NIHIL"))
        else:
            for so in res.so_orders:
                tag = "TERPENUHI"
                if "SEBAGIAN" in so.status:
                    tag = "SEBAGIAN"
                elif "MENUNGGU" in so.status:
                    tag = "MENUNGGU_STOK"

                self.tree.insert(
                    "",
                    tk.END,
                    values=(
                        so.queue_no,
                        so.sales_order,
                        so.created_date,
                        so.sales_name or so.sales_resp,
                        so.gl,
                        so.asm,
                        so.mgr,
                        f"{so.qty_order:,.0f}",
                        f"{so.allocated_qty:,.0f}",
                        so.status
                    ),
                    tags=(tag,)
                )

        self.search_status_lbl.configure(text=f"Analisis untuk item '{res.item_code}' selesai ditampilkan.", fg="#166534")
        self.btn_copy.configure(state=tk.NORMAL)

    def _copy_to_clipboard(self):
        if self.current_result:
            self.win.clipboard_clear()
            self.win.clipboard_append(self.current_result.formatted_report)
            self.copy_toast_lbl.configure(text="[TERSALIN] Ringkasan lengkap berhasil disalin!")
            self.win.after(3000, lambda: self.copy_toast_lbl.configure(text=""))


class RecheckSPWindow:
    """
    Jendela Recheck SP Kemarin & Validasi Posting:
    - Verifikasi SP kemarin yang terlewat diposting di Dynamics AX.
    - Klasifikasi otomatis:
      * [SIAP POSTING] (Terlewat): Stok fisik gudang Makassar WH 56 cukup.
      * [POTENSI SELISIH]: Stok sistem ada tapi fisik di rak kosong / macet kronis.
      * [MENUNGGU STOK]: Stok memang kosong/indent.
    - Panduan Delimit 14 Item per batch siap copy-paste langsung ke filter Dynamics AX.
    - Export langsung dan akses 1-klik ke file Excel RECHECK_SP_<TANGGAL>.xlsx.
    """
    def __init__(self, parent: tk.Tk, initial_dir: str):
        self.parent = parent
        self.work_dir = initial_dir
        self.current_result: Optional[RecheckSPResult] = None
        self.is_running = False
        self.result_queue = queue.Queue()
        self.icons = IconLibrary()

        self.win = tk.Toplevel(parent)
        self.win.title("Daily Data - Recheck SP Kemarin & Validasi Posting")
        self.win.geometry("1100x780")
        self.win.minsize(950, 640)
        self.win.transient(parent)

        self._build_ui()
        self._check_queue()
        self._start_recheck()

    def _build_ui(self):
        # 1. HEADER
        header_frame = tk.Frame(self.win, bg="#1e1b4b", padx=20, pady=12)
        header_frame.pack(fill=tk.X)

        tk.Label(
            header_frame,
            text="RECHECK SP KEMARIN & VALIDASI POSTING",
            font=("Helvetica", 14, "bold"),
            fg="#ffffff",
            bg="#1e1b4b"
        ).pack(anchor="w")

        tk.Label(
            header_frame,
            text="Deteksi Keterlewatan Posting SP Kemarin • Cek Stok Fisik WH 56 Makassar • Panduan Batch Delimit 14 Item Axapta",
            font=("Helvetica", 9),
            fg="#c7d2fe",
            bg="#1e1b4b"
        ).pack(anchor="w")

        # 2. CONTROL BAR
        ctrl_frame = ttk.LabelFrame(self.win, text=" Kontrol & Folder Operasional ", padding=10)
        ctrl_frame.pack(fill=tk.X, padx=15, pady=(8, 4))

        row1 = tk.Frame(ctrl_frame)
        row1.pack(fill=tk.X)

        tk.Label(row1, text="Folder Target:", font=("Helvetica", 9, "bold")).pack(side=tk.LEFT, padx=(0, 6))

        folders = [os.path.basename(d) for d in list_date_folders()]
        curr_name = os.path.basename(self.work_dir)
        self.dir_var = tk.StringVar(value=curr_name if curr_name in folders else (folders[0] if folders else ""))
        self.folder_combo = ttk.Combobox(row1, textvariable=self.dir_var, values=folders, width=22, state="readonly")
        self.folder_combo.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_run = tk.Button(
            row1,
            text=" Jalankan Recheck",
            image=self.icons.get_play_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 9, "bold"),
            bg="#4f46e5",
            fg="white",
            activebackground="#4338ca",
            activeforeground="white",
            padx=12,
            pady=4,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._start_recheck
        )
        self.btn_run.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_open_excel = tk.Button(
            row1,
            text=" Buka File Excel",
            image=self.icons.get_box_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 9, "bold"),
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            padx=12,
            pady=4,
            relief=tk.FLAT,
            cursor="hand2",
            state=tk.DISABLED,
            command=self._open_excel_file
        )
        self.btn_open_excel.pack(side=tk.LEFT, padx=(0, 10))

        self.status_lbl = tk.Label(row1, text="Memuat...", font=("Helvetica", 9, "italic"), fg="#475569")
        self.status_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 3. KPI BADGES FRAME
        badge_frame = tk.Frame(self.win, padx=15, pady=4)
        badge_frame.pack(fill=tk.X)

        self.badge_labels = {}
        badge_defs = [
            ("total", "TOTAL DIPERIKSA", "0 SO (0 Baris)", "#f1f5f9", "#334155", "#cbd5e1"),
            ("ready", "SIAP POSTING (TERLEWAT)", "0 SO (0 Baris)", "#dcfce7", "#15803d", "#86efac"),
            ("diff", "POTENSI SELISIH FISIK", "0 SO (0 Baris)", "#fef3c7", "#b45309", "#fcd34d"),
            ("waiting", "MENUNGGU STOK (INDENT)", "0 SO (0 Baris)", "#f8fafc", "#64748b", "#e2e8f0")
        ]

        for key, title, def_val, bg, fg, bd in badge_defs:
            b_box = tk.Frame(badge_frame, bg=bg, highlightbackground=bd, highlightthickness=1, padx=10, pady=6)
            b_box.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=3)
            tk.Label(b_box, text=title, font=("Helvetica", 8, "bold"), fg=fg, bg=bg).pack(anchor="w")
            lbl_val = tk.Label(b_box, text=def_val, font=("Helvetica", 11, "bold"), fg=fg, bg=bg)
            lbl_val.pack(anchor="w", pady=(2, 0))
            self.badge_labels[key] = lbl_val

        # 4. MAIN NOTEBOOK (TABS)
        self.notebook = ttk.Notebook(self.win)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=(4, 10))

        # Tab 1: Siap Posting
        self.tab_ready = tk.Frame(self.notebook, bg="#f8fafc")
        self.notebook.add(self.tab_ready, text=" 🟢 Siap Posting (Terlewat) ")
        self._build_ready_tab()

        # Tab 2: Delimit Batch 14 Item
        self.tab_delimit = tk.Frame(self.notebook, bg="#f8fafc")
        self.notebook.add(self.tab_delimit, text=" 📋 Panduan Delimit AX (14 Item) ")
        self._build_delimit_tab()

        # Tab 3: Potensi Selisih Fisik
        self.tab_diff = tk.Frame(self.notebook, bg="#f8fafc")
        self.notebook.add(self.tab_diff, text=" ⚠️ Potensi Selisih Fisik ")
        self._build_diff_tab()

        # Tab 4: Semua Transaksi
        self.tab_all = tk.Frame(self.notebook, bg="#f8fafc")
        self.notebook.add(self.tab_all, text=" 📑 Semua Transaksi Evaluasi ")
        self._build_all_tab()

    def _build_ready_tab(self):
        top_bar = tk.Frame(self.tab_ready, bg="#f8fafc", padx=10, pady=6)
        top_bar.pack(fill=tk.X)

        tk.Label(
            top_bar,
            text="Daftar SP kemarin yang stok fisiknya di WH 56 Makassar TERBUKTI ADA & CUKUP, namun belum diposting:",
            font=("Helvetica", 9),
            fg="#166534",
            bg="#f8fafc"
        ).pack(side=tk.LEFT)

        self.ready_toast_lbl = tk.Label(top_bar, text="", font=("Helvetica", 9, "bold"), fg="#15803d", bg="#f8fafc")
        self.ready_toast_lbl.pack(side=tk.RIGHT, padx=10)

        btn_copy_ready = tk.Button(
            top_bar,
            text=" Salin Semua Nomor SO",
            image=self.icons.get_copy_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 8, "bold"),
            bg="#15803d",
            fg="white",
            padx=8,
            pady=3,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._copy_all_ready_sos
        )
        btn_copy_ready.pack(side=tk.RIGHT)

        tree_frame = tk.Frame(self.tab_ready)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        cols = ("no", "so", "sp", "mgr", "asm", "gl", "tgl", "cust", "item", "judul", "pesan", "sisa", "fisik", "sumber")
        self.tree_ready = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")

        col_configs = [
            ("no", "No", 40, "center"),
            ("so", "Sales Order", 130, "center"),
            ("sp", "Surat Pesanan", 110, "center"),
            ("mgr", "Manager", 110, "left"),
            ("asm", "ASM", 110, "left"),
            ("gl", "GL", 110, "left"),
            ("tgl", "Tgl Buat", 120, "center"),
            ("cust", "Customer", 90, "center"),
            ("item", "Item Number", 100, "center"),
            ("judul", "Judul Buku", 240, "left"),
            ("pesan", "Qty", 50, "e"),
            ("sisa", "Sisa", 50, "e"),
            ("fisik", "Fisik 56", 65, "e"),
            ("sumber", "Sumber Alokasi", 150, "left")
        ]

        for cid, head, w, anch in col_configs:
            self.tree_ready.heading(cid, text=head)
            self.tree_ready.column(cid, width=w, anchor=anch)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_ready.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree_ready.xview)
        self.tree_ready.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_ready.pack(fill=tk.BOTH, expand=True)

    def _build_delimit_tab(self):
        top_bar = tk.Frame(self.tab_delimit, bg="#f8fafc", padx=10, pady=6)
        top_bar.pack(fill=tk.X)

        tk.Label(
            top_bar,
            text="Panduan Batch Delimit 14 Item per baris (Langsung copy-paste ke filter Sales Order Dynamics AX):",
            font=("Helvetica", 9),
            fg="#1e3a8a",
            bg="#f8fafc"
        ).pack(side=tk.LEFT)

        self.delimit_container = tk.Frame(self.tab_delimit, bg="#f8fafc")
        self.delimit_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        # Canvas scrollable for batches
        self.delimit_canvas = tk.Canvas(self.delimit_container, bg="#ffffff", highlightthickness=1, highlightbackground="#cbd5e1")
        delimit_vsb = ttk.Scrollbar(self.delimit_container, orient=tk.VERTICAL, command=self.delimit_canvas.yview)
        self.delimit_scroll_frame = tk.Frame(self.delimit_canvas, bg="#ffffff", padx=12, pady=10)

        self.delimit_scroll_frame.bind(
            "<Configure>",
            lambda e: self.delimit_canvas.configure(scrollregion=self.delimit_canvas.bbox("all"))
        )
        self.delimit_canvas_window = self.delimit_canvas.create_window((0, 0), window=self.delimit_scroll_frame, anchor="nw")
        self.delimit_canvas.configure(yscrollcommand=delimit_vsb.set)

        self.delimit_canvas.bind("<Configure>", lambda e: self.delimit_canvas.itemconfig(self.delimit_canvas_window, width=e.width))

        delimit_vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.delimit_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def _build_diff_tab(self):
        top_bar = tk.Frame(self.tab_diff, bg="#f8fafc", padx=10, pady=6)
        top_bar.pack(fill=tk.X)

        tk.Label(
            top_bar,
            text="Daftar item dengan indikasi selisih stok (On-hand sistem ada tapi fisik di rak gudang kosong / pesanan lama mandek):",
            font=("Helvetica", 9),
            fg="#b45309",
            bg="#f8fafc"
        ).pack(side=tk.LEFT)

        btn_copy_diff = tk.Button(
            top_bar,
            text=" Salin Daftar Item (Opname)",
            image=self.icons.get_copy_icon(),
            compound=tk.LEFT,
            font=("Helvetica", 8, "bold"),
            bg="#b45309",
            fg="white",
            padx=8,
            pady=3,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._copy_diff_items
        )
        btn_copy_diff.pack(side=tk.RIGHT)

        tree_frame = tk.Frame(self.tab_diff)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        cols = ("no", "so", "sp", "mgr", "cust", "item", "judul", "sisa", "onh", "fisik", "ket")
        self.tree_diff = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")

        col_configs = [
            ("no", "No", 40, "center"),
            ("so", "Sales Order", 130, "center"),
            ("sp", "Surat Pesanan", 110, "center"),
            ("mgr", "Manager", 110, "left"),
            ("cust", "Customer", 90, "center"),
            ("item", "Item Number", 100, "center"),
            ("judul", "Judul Buku", 240, "left"),
            ("sisa", "Sisa", 50, "e"),
            ("onh", "On-hand", 65, "e"),
            ("fisik", "Fisik 56", 65, "e"),
            ("ket", "Diagnosa Masalah", 280, "left")
        ]

        for cid, head, w, anch in col_configs:
            self.tree_diff.heading(cid, text=head)
            self.tree_diff.column(cid, width=w, anchor=anch)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_diff.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree_diff.xview)
        self.tree_diff.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_diff.pack(fill=tk.BOTH, expand=True)

    def _build_all_tab(self):
        top_bar = tk.Frame(self.tab_all, bg="#f8fafc", padx=10, pady=6)
        top_bar.pack(fill=tk.X)

        tk.Label(
            top_bar,
            text="Seluruh transaksi kemarin yang dievaluasi lengkap beserta diagnosanya:",
            font=("Helvetica", 9),
            fg="#475569",
            bg="#f8fafc"
        ).pack(side=tk.LEFT)

        tree_frame = tk.Frame(self.tab_all)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        cols = ("no", "so", "sp", "mgr", "tgl", "cust", "item", "judul", "sisa", "fisik", "onh", "status", "sumber")
        self.tree_all = ttk.Treeview(tree_frame, columns=cols, show="headings", selectmode="extended")

        col_configs = [
            ("no", "No", 40, "center"),
            ("so", "Sales Order", 130, "center"),
            ("sp", "Surat Pesanan", 110, "center"),
            ("mgr", "Manager", 110, "left"),
            ("tgl", "Tgl Dibuat", 120, "center"),
            ("cust", "Customer", 90, "center"),
            ("item", "Item Number", 100, "center"),
            ("judul", "Judul Buku", 240, "left"),
            ("sisa", "Sisa", 50, "e"),
            ("fisik", "Fisik 56", 65, "e"),
            ("onh", "On-hand", 65, "e"),
            ("status", "Status Diagnosa", 160, "center"),
            ("sumber", "Sumber Alokasi", 150, "left")
        ]

        for cid, head, w, anch in col_configs:
            self.tree_all.heading(cid, text=head)
            self.tree_all.column(cid, width=w, anchor=anch)

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree_all.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree_all.xview)
        self.tree_all.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_all.pack(fill=tk.BOTH, expand=True)

        self.tree_all.tag_configure("READY", background="#f0fdf4")
        self.tree_all.tag_configure("DIFF", background="#fffbeb")
        self.tree_all.tag_configure("WAIT", background="#ffffff")

    def _start_recheck(self):
        if self.is_running:
            return
        selected_folder = self.dir_var.get().strip()
        try:
            target_dir = resolve_work_dir(selected_folder) if selected_folder else PROJECT_ROOT
        except Exception:
            target_dir = PROJECT_ROOT

        self.is_running = True
        self.btn_run.configure(state=tk.DISABLED, bg="#94a3b8")
        self.btn_open_excel.configure(state=tk.DISABLED)
        self.status_lbl.configure(text="Sedang memproses dan menganalisis data stok & SOOUT...", fg="#0284c7")

        threading.Thread(target=self._worker_thread, args=(target_dir,), daemon=True).start()

    def _worker_thread(self, target_dir: str):
        try:
            res = run_recheck_sp(target_dir)
            self.result_queue.put(("SUCCESS", res))
        except Exception as e:
            self.result_queue.put(("ERROR", str(e)))

    def _check_queue(self):
        try:
            while not self.result_queue.empty():
                status, data = self.result_queue.get_nowait()
                self.is_running = False
                self.btn_run.configure(state=tk.NORMAL, bg="#4f46e5")

                if status == "SUCCESS":
                    self.current_result = data
                    self._render_results(data)
                else:
                    self.status_lbl.configure(text=f"Gagal: {data}", fg="#dc2626")
                    messagebox.showerror("Error Recheck", f"Terjadi kesalahan saat recheck:\n{data}")
        except Exception:
            pass
        self.win.after(100, self._check_queue)

    def _render_results(self, res: RecheckSPResult):
        # 1. Update status bar & open excel button
        if res.excel_output_path and os.path.exists(res.excel_output_path):
            self.btn_open_excel.configure(state=tk.NORMAL)
            self.status_lbl.configure(text=f"Analisis selesai! File tersimpan: {os.path.basename(res.excel_output_path)}", fg="#166534")
        else:
            self.status_lbl.configure(text="Analisis selesai.", fg="#166534")

        # 2. Update Badges
        ready_sos_cnt = len({x.sales_order for x in res.ready_to_post})
        diff_sos_cnt = len({x.sales_order for x in res.potential_diff})
        waiting_sos_cnt = len({x.sales_order for x in res.waiting_stock})
        total_sos_cnt = len({x.sales_order for x in res.all_items})

        self.badge_labels["total"].configure(text=f"{total_sos_cnt} SO ({res.total_checked:,} Baris)")
        self.badge_labels["ready"].configure(text=f"{ready_sos_cnt} SO ({len(res.ready_to_post):,} Baris)")
        self.badge_labels["diff"].configure(text=f"{diff_sos_cnt} SO ({len(res.potential_diff):,} Baris)")
        self.badge_labels["waiting"].configure(text=f"{waiting_sos_cnt} SO ({len(res.waiting_stock):,} Baris)")

        # 3. Render Tab 1 (Siap Posting)
        for item in self.tree_ready.get_children():
            self.tree_ready.delete(item)

        for idx, itm in enumerate(res.ready_to_post, 1):
            self.tree_ready.insert(
                "",
                tk.END,
                values=(
                    idx,
                    itm.sales_order,
                    itm.surat_pesanan,
                    itm.mgr,
                    itm.asm,
                    itm.gl,
                    itm.created_date,
                    itm.customer_account,
                    itm.item_number,
                    itm.item_name,
                    f"{itm.qty_ordered:,.0f}",
                    f"{itm.sisa_kirim:,.0f}",
                    f"{itm.stok_fisik_wh56:,.0f}",
                    itm.sumber_alokasi
                )
            )

        # 4. Render Tab 2 (Delimit 14 Item)
        for widget in self.delimit_scroll_frame.winfo_children():
            widget.destroy()

        def add_batch_row(parent, batch_label: str, batch_str: str):
            row_frame = tk.Frame(parent, bg="#ffffff", pady=4)
            row_frame.pack(fill=tk.X)

            tk.Label(row_frame, text=batch_label, font=("Helvetica", 9, "bold"), width=22, anchor="w", bg="#ffffff").pack(side=tk.LEFT)

            entry = ttk.Entry(row_frame, font=("Consolas", 9))
            entry.insert(0, batch_str)
            entry.configure(state="readonly")
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

            btn_copy = tk.Button(
                row_frame,
                text=" Salin",
                image=self.icons.get_copy_icon(),
                compound=tk.LEFT,
                font=("Helvetica", 8, "bold"),
                bg="#0284c7",
                fg="white",
                padx=8,
                pady=2,
                relief=tk.FLAT,
                cursor="hand2"
            )
            btn_copy.pack(side=tk.RIGHT)

            def do_copy(b_val=batch_str, b_btn=btn_copy):
                self.win.clipboard_clear()
                self.win.clipboard_append(b_val)
                b_btn.configure(text=" ✓ Tersalin!", bg="#15803d")
                self.win.after(2000, lambda: b_btn.configure(text=" Salin", bg="#0284c7"))

            btn_copy.configure(command=do_copy)

        if res.delimit_batches_mgr1:
            m1_hdr = tk.Frame(self.delimit_scroll_frame, bg="#e0f2fe", padx=8, pady=4)
            m1_hdr.pack(fill=tk.X, pady=(4, 6))
            tk.Label(m1_hdr, text=f"[ MANAGER 1 - {len({x.sales_order for x in res.ready_to_post if '1' in str(x.mgr)})} SO SIAP POSTING ]", font=("Helvetica", 10, "bold"), fg="#0369a1", bg="#e0f2fe").pack(anchor="w")
            for b_i, b_str in enumerate(res.delimit_batches_mgr1, 1):
                add_batch_row(self.delimit_scroll_frame, f"Batch M1-{b_i} ({len(b_str.split(','))} SO):", b_str)

        if res.delimit_batches_mgr2:
            m2_hdr = tk.Frame(self.delimit_scroll_frame, bg="#fef3c7", padx=8, pady=4)
            m2_hdr.pack(fill=tk.X, pady=(10, 6))
            tk.Label(m2_hdr, text=f"[ MANAGER 2 - {len({x.sales_order for x in res.ready_to_post if '2' in str(x.mgr)})} SO SIAP POSTING ]", font=("Helvetica", 10, "bold"), fg="#b45309", bg="#fef3c7").pack(anchor="w")
            for b_i, b_str in enumerate(res.delimit_batches_mgr2, 1):
                add_batch_row(self.delimit_scroll_frame, f"Batch M2-{b_i} ({len(b_str.split(','))} SO):", b_str)

        if not res.ready_to_post:
            tk.Label(
                self.delimit_scroll_frame,
                text="Tidak ada pesanan yang siap diposting saat ini.",
                font=("Helvetica", 10, "italic"),
                fg="#64748b",
                bg="#ffffff"
            ).pack(pady=20)

        # 5. Render Tab 3 (Potensi Selisih Fisik)
        for item in self.tree_diff.get_children():
            self.tree_diff.delete(item)

        for idx, itm in enumerate(res.potential_diff, 1):
            self.tree_diff.insert(
                "",
                tk.END,
                values=(
                    idx,
                    itm.sales_order,
                    itm.surat_pesanan,
                    itm.mgr,
                    itm.customer_account,
                    itm.item_number,
                    itm.item_name,
                    f"{itm.sisa_kirim:,.0f}",
                    f"{itm.onhand_sistem:,.0f}",
                    f"{itm.stok_fisik_wh56:,.0f}",
                    itm.keterangan
                )
            )

        # 6. Render Tab 4 (Semua Transaksi)
        for item in self.tree_all.get_children():
            self.tree_all.delete(item)

        for idx, itm in enumerate(res.all_items, 1):
            tag = "WAIT"
            if itm.status_diagnosa == "SIAP POSTING (TERLEWAT)":
                tag = "READY"
            elif itm.status_diagnosa == "POTENSI SELISIH FISIK":
                tag = "DIFF"

            self.tree_all.insert(
                "",
                tk.END,
                values=(
                    idx,
                    itm.sales_order,
                    itm.surat_pesanan,
                    itm.mgr,
                    itm.created_date,
                    itm.customer_account,
                    itm.item_number,
                    itm.item_name,
                    f"{itm.sisa_kirim:,.0f}",
                    f"{itm.stok_fisik_wh56:,.0f}",
                    f"{itm.onhand_sistem:,.0f}",
                    itm.status_diagnosa,
                    itm.sumber_alokasi
                ),
                tags=(tag,)
            )

    def _copy_all_ready_sos(self):
        if not self.current_result or not self.current_result.ready_to_post:
            return
        so_list = []
        for x in self.current_result.ready_to_post:
            if x.sales_order not in so_list:
                so_list.append(x.sales_order)
        joined = ",".join(so_list)
        self.win.clipboard_clear()
        self.win.clipboard_append(joined)
        self.ready_toast_lbl.configure(text=f"[TERSALIN] {len(so_list)} nomor SO berhasil disalin!")
        self.win.after(3000, lambda: self.ready_toast_lbl.configure(text=""))

    def _copy_diff_items(self):
        if not self.current_result or not self.current_result.potential_diff:
            return
        lines = ["Item Number\tJudul Buku\tSisa Butuh\tOn-hand Sistem\tStok Fisik WH56\tKeterangan"]
        seen_it = set()
        for x in self.current_result.potential_diff:
            if x.item_number not in seen_it:
                seen_it.add(x.item_number)
                lines.append(f"{x.item_number}\t{x.item_name}\t{x.sisa_kirim:.0f}\t{x.onhand_sistem:.0f}\t{x.stok_fisik_wh56:.0f}\t{x.keterangan}")
        text_out = "\n".join(lines)
        self.win.clipboard_clear()
        self.win.clipboard_append(text_out)
        messagebox.showinfo("Tersalin", f"{len(seen_it)} item selisih berhasil disalin dalam format tabel!")

    def _open_excel_file(self):
        if self.current_result and self.current_result.excel_output_path:
            p = self.current_result.excel_output_path
            if os.path.exists(p):
                try:
                    os.startfile(p)
                except Exception as e:
                    messagebox.showerror("Error", f"Gagal membuka file Excel:\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DailyDataGUI(root)
    root.mainloop()
