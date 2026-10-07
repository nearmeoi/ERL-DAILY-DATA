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
    StageStatus,
    StageMetric,
    DailyDataError,
    FileLockError,
    StageExecutionError
)

class DailyDataGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Daily Data - Sistem Otomasi Pipeline")
        self.root.geometry("880x680")
        self.root.minsize(760, 560)

        # Styling tema
        style = ttk.Style()
        style.theme_use('clam')

        self.log_queue = queue.Queue()
        self.stage_queue = queue.Queue()
        self.is_running = False
        self.stage_widgets = {}

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

        refresh_btn = ttk.Button(config_frame, text="Refresh", command=self._refresh_folders)
        refresh_btn.pack(side=tk.RIGHT, padx=(5, 0))

        browse_btn = ttk.Button(config_frame, text="Pilih Folder...", command=self._browse_dir)
        browse_btn.pack(side=tk.RIGHT)

        # 3. ACTION CONTROLS (RUN BUTTON)
        action_frame = tk.Frame(self.root, padx=15, pady=6)
        action_frame.pack(fill=tk.X)

        # Tombol RUN bergaya modern minimalis
        self.run_btn = tk.Button(
            action_frame,
            text="▶  RUN",
            font=("Helvetica", 11, "bold"),
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            padx=28,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._start_pipeline
        )
        self.run_btn.pack(side=tk.LEFT)

        self.progress_bar = ttk.Progressbar(action_frame, mode='indeterminate')
        self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(15, 0))

        # 4. TIMELINE TRACKER (Visual Timeline seperti Video Editor)
        timeline_frame = ttk.LabelFrame(self.root, text=" Timeline Pipeline Eksekusi ", padding=8)
        timeline_frame.pack(fill=tk.X, padx=15, pady=(4, 8))

        timeline_badges_frame = tk.Frame(timeline_frame)
        timeline_badges_frame.pack(fill=tk.X)

        # Definisi tahapan timeline
        self.stage_defs = [
            (1, "01 Init"),
            (2, "02 Stok"),
            (3, "03 Trn"),
            (4, "04 SOOUT"),
            (5, "05 POOUT"),
            (6, "06 Format"),
            (7, "07 Bukmas"),
            (8, "08 DO Sum"),
            (9, "09 Sync")
        ]

        for idx, short_name in self.stage_defs:
            badge_lbl = tk.Label(
                timeline_badges_frame,
                text=f"○ {short_name}",
                font=("Consolas", 8, "bold"),
                bg="#e2e8f0",
                fg="#475569",
                padx=6,
                pady=4,
                relief=tk.GROOVE
            )
            badge_lbl.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
            self.stage_widgets[idx] = badge_lbl

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
                    widget.configure(text=f"⏳ {short_name}", bg="#0284c7", fg="#ffffff")
                elif stg.status == StageStatus.SUCCESS:
                    widget.configure(text=f"✓ {short_name}", bg="#15803d", fg="#ffffff")
                elif stg.status == StageStatus.SKIPPED:
                    widget.configure(text=f"⊘ {short_name}", bg="#94a3b8", fg="#ffffff")
                elif stg.status == StageStatus.FAILED:
                    widget.configure(text=f"✕ {short_name}", bg="#dc2626", fg="#ffffff")
                else:
                    widget.configure(text=f"○ {short_name}", bg="#e2e8f0", fg="#475569")

        self.root.after(80, self._check_queues)

    def _reset_timeline(self):
        for idx, short_name in self.stage_defs:
            w = self.stage_widgets.get(idx)
            if w:
                w.configure(text=f"○ {short_name}", bg="#e2e8f0", fg="#475569")

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

        self.is_running = True
        self.run_btn.configure(state=tk.DISABLED, bg="#94a3b8")
        self.progress_bar.start(10)
        self.status_var.set("Menjalankan pipeline Daily Data... Harap tunggu.")
        self.log_text.delete(1.0, tk.END)
        self._reset_timeline()

        threading.Thread(target=self._run_thread, args=(target_dir,), daemon=True).start()

    def _run_thread(self, target_dir: str):
        try:
            timeline = run_automation(
                work_dir=target_dir,
                progress_callback=self._append_log,
                stage_callback=self._stage_update_callback
            )
            self._append_log("\n✓ SELURUH TAHAPAN PIPELINE DAILY DATA SELESAI DENGAN SUKSES!")
            self.root.after(0, lambda: messagebox.showinfo(
                "Sukses",
                "Pipeline Daily Data berhasil diselesaikan 100% tanpa error!"
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
            text="⚠️ TERJADI KESALAHAN SAAT EKSEKUSI",
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
            tk.Label(body, text="💡 Solusi Praktis:", font=("Helvetica", 9, "bold")).pack(anchor="w")
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
        self.progress_bar.stop()

if __name__ == "__main__":
    root = tk.Tk()
    app = DailyDataGUI(root)
    root.mainloop()
