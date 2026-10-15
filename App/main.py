import os
import sys
import glob
import re
import shutil
import time
import traceback
from enum import Enum
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Any, Callable

import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from process_do_summary import process_do_summary_file, find_do_summary_for_date
from recheck_sp import run_recheck_sp, RecheckSPResult

# ==============================================================================
# 1. KONFIGURASI GLOBAL & LOKASI DIREKTORI
# ==============================================================================
if getattr(sys, 'frozen', False):
    PROJECT_ROOT = os.path.dirname(os.path.abspath(sys.executable))
    APP_DIR = os.path.join(PROJECT_ROOT, "App")
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    PROJECT_ROOT = os.path.dirname(APP_DIR)
EXAMPLES_DIR = os.path.join(APP_DIR, "examples")

# ==============================================================================
# 2. STRUKTUR DATA & EXCEPTION HIERARCHY (ENTERPRISE ERROR HANDLING)
# ==============================================================================
class StageStatus(str, Enum):
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    SUCCESS = "SUCCESS"
    SKIPPED = "SKIPPED"
    FAILED = "FAILED"

class PipelineMode(str, Enum):
    ALL = "ALL"                      # Tahap 1 - 10 (Full Pipeline)
    POSISI_STOK = "POSISI_STOK"      # Tahap 1, 2, 3, 4, 5, 6, 9, 10 (Pagi: STOK -> Format Cabang + Recheck SP)
    RECHECK_SP = "RECHECK_SP"        # Tahap 1, 9 (Recheck SP Kemarin Saja)
    BUKU_MASUK = "BUKU_MASUK"        # Tahap 1, 7, 10 (Siang: Buku Masuk & FIFO Saja)
    DO_SUMMARY = "DO_SUMMARY"        # Tahap 1, 8, 10 (Sore: DO Summary External Saja)
    CUSTOM = "CUSTOM"                # Sesuai tahapan yang dipilih user

MODE_DESCRIPTIONS = {
    PipelineMode.ALL: "Semua Tahap (1 - 9) [Full Pipeline]",
    PipelineMode.POSISI_STOK: "Posisi Stok Cabang + Recheck SP (Tahap 1-6, 9) [Pagi]",
    PipelineMode.RECHECK_SP: "Recheck SP Kemarin (Validasi Posting Saja)",
    PipelineMode.BUKU_MASUK: "Buku Masuk & FIFO (Tahap 7 Saja) [Siang]",
    PipelineMode.DO_SUMMARY: "DO Summary External (Tahap 8 Saja) [Sore]",
    PipelineMode.CUSTOM: "Kustom (Pilih Tahap Sendiri)"
}

MODE_DEFAULT_STAGES = {
    PipelineMode.ALL: [1, 2, 3, 4, 5, 6, 7, 8, 9],
    PipelineMode.POSISI_STOK: [1, 2, 3, 4, 5, 6, 9],
    PipelineMode.RECHECK_SP: [1, 9],
    PipelineMode.BUKU_MASUK: [1, 7],
    PipelineMode.DO_SUMMARY: [1, 8],
}

class DailyDataError(Exception):
    """Base class untuk seluruh custom exception aplikasi Daily Data."""
    def __init__(self, message: str, resolution_tip: str = ""):
        super().__init__(message)
        self.message = message
        self.resolution_tip = resolution_tip

class FileLockError(DailyDataError):
    """Terjadi ketika file Excel sedang dibuka/dikunci oleh Microsoft Excel."""
    def __init__(self, file_path: str):
        filename = os.path.basename(file_path)
        msg = f"File '{filename}' sedang dibuka atau dikunci oleh program lain (Microsoft Excel)."
        tip = f"Harap TUTUP file '{filename}' di Microsoft Excel terlebih dahulu, lalu klik RUN kembali."
        super().__init__(msg, tip)
        self.file_path = file_path

class MandatoryFileNotFoundError(DailyDataError):
    """File utama yang wajib ada tidak ditemukan di direktori kerja."""
    def __init__(self, pattern: str, target_dir: str):
        msg = f"File wajib dengan pola '{pattern}' tidak ditemukan di: {target_dir}"
        tip = "Pastikan file harian telah disalin ke folder tanggal yang dipilih sebelum menjalankan proses."
        super().__init__(msg, tip)

class StageExecutionError(DailyDataError):
    """Error wrapper saat sebuah tahapan proses (stage) mengalami kegagalan."""
    def __init__(self, stage_idx: int, stage_name: str, original_error: Exception, tb_str: str):
        msg = f"Kegagalan pada Tahap [{stage_idx}] {stage_name}: {str(original_error)}"
        tip = getattr(original_error, "resolution_tip", "Periksa rincian log kesalahan di bawah ini atau hubungi tim teknis.")
        super().__init__(msg, tip)
        self.stage_idx = stage_idx
        self.stage_name = stage_name
        self.original_error = original_error
        self.traceback_str = tb_str

@dataclass
class StageMetric:
    index: int
    id: str
    name: str
    status: StageStatus = StageStatus.PENDING
    start_time: Optional[float] = None
    end_time: Optional[float] = None
    duration_sec: float = 0.0
    summary: str = ""
    error_msg: Optional[str] = None
    error_tip: Optional[str] = None
    error_tb: Optional[str] = None

class ExecutionTimeline:
    """
    Timeline Pelacak Eksekusi (analog visual video editor timeline):
    Mencatat status, durasi per milidetik, metrik, dan lokasi error per tahap.
    """
    def __init__(self):
        self.stages: List[StageMetric] = []
        self._current_index: int = 0

    def register_stage(self, stage_id: str, name: str) -> int:
        idx = len(self.stages) + 1
        metric = StageMetric(index=idx, id=stage_id, name=name)
        self.stages.append(metric)
        return idx

    def get_stage(self, index: int) -> StageMetric:
        return self.stages[index - 1]

    def start_stage(self, index: int):
        stage = self.get_stage(index)
        stage.status = StageStatus.RUNNING
        stage.start_time = time.perf_counter()

    def complete_stage(self, index: int, summary: str = ""):
        stage = self.get_stage(index)
        stage.status = StageStatus.SUCCESS
        stage.end_time = time.perf_counter()
        if stage.start_time:
            stage.duration_sec = stage.end_time - stage.start_time
        stage.summary = summary

    def skip_stage(self, index: int, reason: str = ""):
        stage = self.get_stage(index)
        stage.status = StageStatus.SKIPPED
        stage.summary = reason or "Dilewati (Kondisi tidak terpenuhi)"

    def fail_stage(self, index: int, error: Exception, tb: str):
        stage = self.get_stage(index)
        stage.status = StageStatus.FAILED
        stage.end_time = time.perf_counter()
        if stage.start_time:
            stage.duration_sec = stage.end_time - stage.start_time
        stage.error_msg = str(error)
        stage.error_tip = getattr(error, "resolution_tip", "")
        stage.error_tb = tb

    def get_summary_text(self) -> str:
        lines = [
            "================================================================================",
            "                   RINGKASAN TIMELINE EKSEKUSI DAILY DATA                       ",
            "================================================================================"
        ]
        total_duration = 0.0
        success_count = 0
        fail_count = 0
        skip_count = 0

        for s in self.stages:
            total_duration += s.duration_sec
            if s.status == StageStatus.SUCCESS:
                success_count += 1
                badge = "[SUCCESS]"
            elif s.status == StageStatus.FAILED:
                fail_count += 1
                badge = "[FAILED ]"
            elif s.status == StageStatus.SKIPPED:
                skip_count += 1
                badge = "[SKIPPED]"
            else:
                badge = "[PENDING]"

            dur_str = f"{s.duration_sec:6.2f}s" if s.duration_sec > 0 else "     - "
            detail = f" -> {s.summary}" if s.summary else ""
            lines.append(f" {badge} Tahap {s.index:02d}: {s.name:<32s} | {dur_str}{detail}")

        lines.append("--------------------------------------------------------------------------------")
        status_line = f" Hasil: {success_count} Berhasil, {fail_count} Gagal, {skip_count} Dilewati | Total Durasi: {total_duration:.2f} detik"
        lines.append(status_line)
        lines.append("================================================================================")
        return "\n".join(lines)

# ==============================================================================
# 3. HELPER UTILITIES & ROBUST DATA PARSERS
# ==============================================================================
def log(msg: str):
    timestamp = datetime.now().strftime("[%H:%M:%S]")
    try:
        print(f"{timestamp} {msg}", flush=True)
    except UnicodeEncodeError:
        safe_msg = msg.encode("ascii", errors="replace").decode("ascii")
        print(f"{timestamp} {safe_msg}", flush=True)

def find_file(pattern: str, search_dirs: Optional[List[str]] = None) -> Optional[str]:
    if search_dirs is None:
        search_dirs = [os.getcwd(), PROJECT_ROOT, "D:\\", "D:\\ANALISA STOK"]
    for d in search_dirs:
        if not d or not os.path.exists(d):
            continue
        matches = glob.glob(os.path.join(d, pattern))
        if matches:
            matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            return matches[0]
    return None

def extract_date_str(filename: Optional[str]) -> str:
    if not filename:
        return datetime.now().strftime("%d %b %Y").upper()
    base = os.path.basename(filename)
    m = re.search(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})', base)
    if m:
        parts = m.group(1).split()
        day = parts[0].zfill(2)
        month = parts[1].upper()
        year = parts[2]
        return f"{day} {month} {year}"
    return datetime.now().strftime("%d %b %Y").upper()

def normalize_item(val: Any) -> str:
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def normalize_wh(val: Any) -> str:
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def to_num(series: pd.Series) -> pd.Series:
    """
    Konversi angka aman enterprise:
    Membedakan nilai float/int numerik AX dengan format teks ribuan Indonesia.
    """
    s = series.astype(str).str.strip()
    direct = pd.to_numeric(s, errors="coerce")
    need_fix = direct.isna() & s.str.contains(r"[.,]", regex=True)
    if need_fix.any():
        fixed = pd.to_numeric(
            s[need_fix].str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
            errors="coerce"
        )
        direct.loc[need_fix] = fixed
    return direct.fillna(0)

def list_date_folders(root: str = PROJECT_ROOT) -> List[str]:
    """Daftar subfolder harian yang memuat file FORMAT POSISISTOK."""
    out = []
    if os.path.exists(root):
        for d in glob.glob(os.path.join(root, "*")):
            if os.path.isdir(d) and glob.glob(os.path.join(d, "FORMAT POSISISTOK*.xlsx")):
                out.append(d)
    out.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    return out

def resolve_work_dir(name_or_path: Optional[str]) -> str:
    """
    Resolusi fleksibel nama folder harian (misal '07 OKT 2026' atau path lengkap).
    """
    if not name_or_path:
        raise ValueError("Nama folder target tidak boleh kosong.")
    p = name_or_path.strip().strip('"')
    if os.path.isabs(p) and os.path.isdir(p):
        return p
    cand = os.path.join(PROJECT_ROOT, p)
    if os.path.isdir(cand):
        return cand
    key = " ".join(p.upper().split())
    for d in glob.glob(os.path.join(PROJECT_ROOT, "*")):
        if os.path.isdir(d):
            base = " ".join(os.path.basename(d).upper().split())
            if base == key or base.lstrip("0") == key.lstrip("0"):
                return d
    raise FileNotFoundError(f"Folder harian '{name_or_path}' tidak ditemukan di {PROJECT_ROOT}")

def write_pivot_sheet(
    wb: openpyxl.Workbook,
    header_value: str,
    pivot_dict: Dict[str, float],
    sheet_name: str = "Sheet2",
    page_header: Optional[tuple] = None
):
    """
    Menulis hasil pivot tabel standar:
    Diletakkan di tab pertama, baris 3 header, tanpa item kosong, dan Grand Total di akhir.
    Jika page_header diberikan (misal ('Warehouse', '56')), baris 1 akan memuat filter.
    """
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 0)
    if page_header:
        ws.cell(row=1, column=1, value=page_header[0])
        ws.cell(row=1, column=2, value=page_header[1])
    ws.cell(row=3, column=1, value="Row Labels")
    ws.cell(row=3, column=2, value=header_value)
    r = 4
    total = 0.0
    for item, q in sorted((k, v) for k, v in pivot_dict.items() if k):
        ws.cell(row=r, column=1, value=item)
        c = ws.cell(row=r, column=2, value=int(round(q)))
        c.number_format = "General"
        total += q
        r += 1
    ws.cell(row=r, column=1, value="Grand Total")
    c_tot = ws.cell(row=r, column=2, value=int(round(total)))
    c_tot.number_format = "General"
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 24
    return ws

def repair_external_links(xlsx_path: str) -> bool:
    """
    Memperbaiki ketidaksesuaian relasi externalLink openpyxl:
    Menghilangkan pop-up Excel 'Repaired Records: External formula reference...'.
    """
    if not xlsx_path or not os.path.exists(xlsx_path):
        return False
    import zipfile
    import xml.etree.ElementTree as ET
    import tempfile
    try:
        temp_fd, temp_path = tempfile.mkstemp(suffix='.xlsx')
        os.close(temp_fd)
        modified = False
        with zipfile.ZipFile(xlsx_path, 'r') as zin, zipfile.ZipFile(temp_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            ext_xmls = [n for n in zin.namelist() if n.startswith('xl/externalLinks/') and n.endswith('.xml') and not n.startswith('xl/externalLinks/_rels/')]
            rels_map = {}
            for xname in ext_xmls:
                bname = os.path.basename(xname)
                rname = 'xl/externalLinks/_rels/' + bname + '.rels'
                if rname in zin.namelist():
                    xml_root = ET.fromstring(zin.read(xname))
                    book = xml_root.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}externalBook')
                    target_rid = book.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id') if book is not None else None
                    rels_root = ET.fromstring(zin.read(rname))
                    existing_ids = [r.attrib.get('Id') for r in rels_root]
                    if target_rid and target_rid not in existing_ids:
                        first_rel = rels_root[0] if len(rels_root) > 0 else None
                        if first_rel is not None:
                            new_rel = ET.SubElement(rels_root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship')
                            new_rel.attrib['Id'] = target_rid
                            new_rel.attrib['Type'] = first_rel.attrib.get('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath')
                            new_rel.attrib['Target'] = first_rel.attrib.get('Target', '')
                            new_rel.attrib['TargetMode'] = first_rel.attrib.get('TargetMode', 'External')
                            rels_map[rname] = ET.tostring(rels_root, encoding='utf-8', xml_declaration=True)
                            modified = True
            for item in zin.infolist():
                if item.filename in rels_map:
                    zout.writestr(item.filename, rels_map[item.filename])
                else:
                    zout.writestr(item, zin.read(item.filename))
        if modified:
            shutil.move(temp_path, xlsx_path)
            return True
        else:
            if os.path.exists(temp_path):
                os.remove(temp_path)
            return False
    except Exception:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass
        return False

def check_file_locked(file_path: Optional[str]) -> bool:
    """
    Mengecek secara non-destruktif apakah sebuah file sedang dibuka/dikunci
    oleh program lain seperti Microsoft Excel.
    """
    if not file_path or not os.path.exists(file_path):
        return False
    try:
        with open(file_path, "r+b"):
            pass
        return False
    except (PermissionError, IOError):
        return True

def safe_save_workbook(wb: openpyxl.Workbook, path: str):
    """
    Menyimpan file openpyxl dengan penanganan ramah jika file sedang terkunci di Excel.
    """
    for link in getattr(wb, "_external_links", []):
        if hasattr(link, "externalBook") and hasattr(link, "file_link"):
            if link.externalBook and link.file_link:
                link.file_link.id = link.externalBook.id
    try:
        wb.save(path)
    except PermissionError as pe:
        raise FileLockError(path) from pe
    repair_external_links(path)

def build_sales_hierarchy(su_path: str) -> Dict[str, Dict[str, str]]:
    df_su = pd.read_excel(su_path, sheet_name=0, dtype=str)
    cols_lower = [c.lower() for c in df_su.columns]
    if any("kodsal" in c for c in cols_lower):
        col_kod = [c for c in df_su.columns if "kodsal" in c.lower()][0]
        col_gl = [c for c in df_su.columns if "gl" in c.lower()][0]
        col_asm = [c for c in df_su.columns if "asm" in c.lower()][0]
        col_mgr = [c for c in df_su.columns if "mgr" in c.lower()][0]
        su_db = {}
        for _, row in df_su.iterrows():
            kod = str(row[col_kod]).strip()
            su_db[kod] = {
                "GL": str(row[col_gl]).strip() if pd.notna(row[col_gl]) else "",
                "ASM": str(row[col_asm]).strip() if pd.notna(row[col_asm]) else "",
                "MGR": str(row[col_mgr]).strip() if pd.notna(row[col_mgr]) else ""
            }
        return su_db

    unit_col = [c for c in df_su.columns if "sales unit" in c.lower()][0]
    parent_col = [c for c in df_su.columns if "parent" in c.lower()][0]
    desc_col = [c for c in df_su.columns if "desc" in c.lower()][0]
    emp_col = [c for c in df_su.columns if "emp" in c.lower() or "sales responsible" in c.lower()][0]

    unit_to_desc = dict(zip(df_su[unit_col].str.strip(), df_su[desc_col].str.strip()))
    unit_to_parent = dict(zip(df_su[unit_col].str.strip(), df_su[parent_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()))
    emp_to_unit = dict(zip(df_su[emp_col].str.strip(), df_su[unit_col].str.strip()))

    su_db = {}
    for emp, unit_id in emp_to_unit.items():
        if not emp or pd.isna(emp) or emp == "nan":
            continue
        gl, asm, mgr = "", "", ""
        curr = unit_id
        visited = set()
        while curr and curr in unit_to_desc and curr not in visited:
            visited.add(curr)
            desc = unit_to_desc[curr]
            d_clean = desc.replace("(", "").replace(")", "").strip()
            if "MGR" in desc.upper() or "MANAGER" in desc.upper():
                if not mgr:
                    mgr = "MGR 1 MAKASSAR" if "1" in desc else ("MGR 2 MAKASSAR" if "2" in desc else d_clean)
            elif "ASM" in desc.upper():
                if not asm:
                    asm = d_clean
            elif "GL" in desc.upper() or "REP" in desc.upper():
                if not gl:
                    gl = d_clean
            curr = unit_to_parent.get(curr)
        su_db[emp] = {"GL": gl, "ASM": asm, "MGR": mgr}
    return su_db

# ==============================================================================
# 4. FUNGSI-FUNGSI TAHAPAN PIPELINE (MODULAR STAGES)
# ==============================================================================
@dataclass
class InputBundle:
    work_dir: str
    date_str: str
    format_cabang_path: Optional[str]
    stok_path: Optional[str]
    trn_path: Optional[str]
    soout_path: Optional[str]
    poout_path: Optional[str]
    su_path: Optional[str]
    buku_masuk_path: Optional[str]
    buku_masuk_paths: List[str] = field(default_factory=list)
    do_summary_path: Optional[str] = None

def stage_1_discover_inputs(
    work_dir: str,
    report_fn: Callable[[str], None],
    target_stages: Optional[List[int]] = None
) -> InputBundle:
    if target_stages is None:
        target_stages = [1, 2, 3, 4, 5, 6, 7, 8, 9]

    report_fn(f"Menelaah file input pada folder: {work_dir}")
    format_cabang_path = find_file("FORMAT POSISISTOK*.xlsx", [work_dir])
    if 6 in target_stages and not format_cabang_path:
        raise MandatoryFileNotFoundError("FORMAT POSISISTOK*.xlsx", work_dir)

    trn_path = find_file("TRN*.xlsx", [work_dir, "D:\\ANALISA STOK"])
    if 3 in target_stages and not trn_path:
        raise MandatoryFileNotFoundError("TRN*.xlsx", work_dir)

    soout_path = find_file("SOOUT*.xlsx", [work_dir])
    if (4 in target_stages or 7 in target_stages) and not soout_path:
        raise MandatoryFileNotFoundError("SOOUT*.xlsx", work_dir)

    poout_path = find_file("POOUT*.xlsx", [work_dir])
    if 5 in target_stages and not poout_path:
        raise MandatoryFileNotFoundError("POOUT*.xlsx", work_dir)

    stok_path = find_file("STOK*.xlsx", [work_dir, "D:\\ANALISA STOK", "D:\\"])
    if 2 in target_stages and not stok_path:
        raise MandatoryFileNotFoundError("STOK*.xlsx", work_dir)

    ref_dir = os.path.join(PROJECT_ROOT, "REFERENSI")
    su_path = find_file("SALESUNIT2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\", APP_DIR]) or \
              find_file("SU2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\", APP_DIR])

    # Kumpulkan seluruh file Buku Masuk (bisa multi-file jika ada pengiriman pagi & siang)
    bm_files = []
    for pat in ["buku masuk*.xlsx", "BUKMAS*.xlsx"]:
        for match in glob.glob(os.path.join(work_dir, pat)):
            if "PRINT" not in os.path.basename(match).upper() and match not in bm_files:
                bm_files.append(match)
    bm_files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    buku_masuk_path = bm_files[0] if bm_files else None

    date_str = extract_date_str(format_cabang_path or buku_masuk_path or soout_path or work_dir)
    do_summary_path = None
    if 8 in target_stages:
        do_summary_path = find_do_summary_for_date(work_dir, target_date_str=date_str, report_fn=report_fn)

    report_fn(f"  - Format Cabang : {os.path.basename(format_cabang_path) if format_cabang_path else '[TIDAK DIGUNAKAN]'}")
    report_fn(f"  - STOK          : {os.path.basename(stok_path) if stok_path else '[TIDAK DIGUNAKAN / TIDAK ADA]'}")
    report_fn(f"  - TRN           : {os.path.basename(trn_path) if trn_path else '[TIDAK DIGUNAKAN / TIDAK ADA]'}")
    report_fn(f"  - SOOUT         : {os.path.basename(soout_path) if soout_path else '[TIDAK DIGUNAKAN / TIDAK ADA]'}")
    report_fn(f"  - POOUT         : {os.path.basename(poout_path) if poout_path else '[TIDAK DIGUNAKAN / TIDAK ADA]'}")
    report_fn(f"  - Master Sales  : {os.path.basename(su_path) if su_path else '[TIDAK DITEMUKAN]'}")
    if bm_files:
        bm_names = ", ".join([os.path.basename(f) for f in bm_files])
        report_fn(f"  - Buku Masuk    : {bm_names} ({len(bm_files)} file)")
    else:
        report_fn(f"  - Buku Masuk    : [TIDAK ADA - SKIP]")
    if 8 in target_stages:
        report_fn(f"  - DO Summary    : {os.path.basename(do_summary_path) if do_summary_path else '[BELUM ADA DI FOLDER]'}")
    report_fn(f"  - Tanggal Opr   : {date_str}")

    locked_warns = []
    if 6 in target_stages and format_cabang_path and check_file_locked(format_cabang_path):
        locked_warns.append(os.path.basename(format_cabang_path))
    if 3 in target_stages and trn_path and check_file_locked(trn_path):
        locked_warns.append(os.path.basename(trn_path))
    if 4 in target_stages and soout_path and check_file_locked(soout_path):
        locked_warns.append(os.path.basename(soout_path))
    if 5 in target_stages and poout_path and check_file_locked(poout_path):
        locked_warns.append(os.path.basename(poout_path))
    if 2 in target_stages and stok_path and check_file_locked(stok_path):
        locked_warns.append(os.path.basename(stok_path))

    if locked_warns:
        report_fn(f"  [PERINGATAN KUNCI FILE] File berikut sedang dibuka di Excel: {', '.join(locked_warns)}")

    return InputBundle(
        work_dir=work_dir,
        date_str=date_str,
        format_cabang_path=format_cabang_path,
        stok_path=stok_path,
        trn_path=trn_path,
        soout_path=soout_path,
        poout_path=poout_path,
        su_path=su_path,
        buku_masuk_path=buku_masuk_path,
        buku_masuk_paths=bm_files,
        do_summary_path=do_summary_path
    )

@dataclass
class StokBundle:
    pst: Dict[str, float]
    nas: Dict[str, float]
    mks: Dict[str, float]

def stage_2_process_stok(stok_path: Optional[str], report_fn: Callable[[str], None]) -> StokBundle:
    if not stok_path or not os.path.exists(stok_path):
        report_fn("File STOK tidak ditemukan. Nilai STOK diatur default 0.")
        return StokBundle(pst={}, nas={}, mks={})

    report_fn(f"Membaca file STOK: {os.path.basename(stok_path)}...")
    df_stok = pd.read_excel(stok_path, sheet_name="Sheet1", dtype=str)
    wh_col = [c for c in df_stok.columns if "warehouse" in c.lower()][0]
    it_col = [c for c in df_stok.columns if "item number" in c.lower()][0]
    qty_col = [c for c in df_stok.columns if "available physical" in c.lower()][0]

    df_stok[qty_col] = to_num(df_stok[qty_col])
    df_stok['wh_clean'] = df_stok[wh_col].apply(normalize_wh)
    df_stok['item_clean'] = df_stok[it_col].apply(normalize_item)

    # Filter baris yang valid (item dan warehouse tidak kosong)
    df_stok = df_stok[(df_stok['item_clean'] != "") & (df_stok['wh_clean'] != "")].copy()

    df_pst = df_stok[df_stok['wh_clean'] == '50'].groupby('item_clean')[qty_col].sum()
    stok_pst = {k: v for k, v in df_pst.to_dict().items() if k}

    df_mks = df_stok[df_stok['wh_clean'] == '56'].groupby('item_clean')[qty_col].sum()
    stok_mks = {k: v for k, v in df_mks.to_dict().items() if k}

    # Rule Baku: Nasional = Semua kecuali 50, 56, 723
    df_nas = df_stok[~df_stok['wh_clean'].isin(['50', '56', '723'])].groupby('item_clean')[qty_col].sum()
    stok_nas = {k: v for k, v in df_nas.to_dict().items() if k}

    wh_available = sorted(df_stok['wh_clean'].unique())
    report_fn(f"Gudang terdeteksi: {', '.join(wh_available)}")
    report_fn(f"Agregasi selesai: {len(stok_pst):,} item Pusat, {len(stok_nas):,} item Nasional, {len(stok_mks):,} item Makassar")
    if len(stok_nas) == 0:
        report_fn("  [PERHATIAN] Data STOK Nasional tidak tersedia pada file STOK (hanya Pusat & Makassar).")
        report_fn("  Kolom STOK_NAS pada Format Cabang akan dipertahankan dari data sebelumnya (tidak di-overwrite 0).")

    # Pembuatan Sheet2 Pivot STOK (standar konsisten sama dengan TRN, SOOUT, POOUT)
    wb_stok = openpyxl.load_workbook(stok_path)
    default_wh = "56" if "56" in wh_available else ("(All)" if len(wh_available) > 1 else wh_available[0])
    fallback_dict = stok_mks if "56" in wh_available else {k: v for k, v in df_stok.groupby('item_clean')[qty_col].sum().to_dict().items() if k}
    write_pivot_sheet(wb_stok, "Sum of Available physical", fallback_dict, page_header=("Warehouse", default_wh))
    safe_save_workbook(wb_stok, stok_path)

    try:
        from native_pivot import create_native_excel_pivot
        create_native_excel_pivot(
            xlsx_path=stok_path,
            table_name="MyGrid",
            target_sheet="Sheet2",
            row_field="Item number",
            val_field="Available physical",
            page_field="Warehouse",
            page_item=default_wh,
            style_name="PivotStyleLight16",
            log_func=report_fn
        )
    except Exception as e:
        report_fn(f"  [PIVOT NOTE] Menggunakan tabel statis cadangan: {e}")

    report_fn(f"Sheet2 Pivot STOK diperbarui ({len(fallback_dict):,} item unik)")
    return StokBundle(pst=stok_pst, nas=stok_nas, mks=stok_mks)

@dataclass
class TrnBundle:
    pivot: Dict[str, float]

def stage_3_process_trn(trn_path: Optional[str], report_fn: Callable[[str], None]) -> TrnBundle:
    if not trn_path or not os.path.exists(trn_path):
        report_fn("File TRN tidak ditemukan.")
        return TrnBundle(pivot={})

    report_fn(f"Membaca file TRN: {os.path.basename(trn_path)}...")
    df_trn = pd.read_excel(trn_path, sheet_name="Sheet1", dtype=str)
    qdate_col = [c for c in df_trn.columns if "quarantine date" in c.lower()][0]
    it_col = [c for c in df_trn.columns if "item number" in c.lower()][0]
    qty_col = [c for c in df_trn.columns if "quantity" in c.lower()][0]

    qdate_parsed = pd.to_datetime(df_trn[qdate_col], errors="coerce")
    mask_2026 = (qdate_parsed.dt.year == 2026) | (qdate_parsed.isna() & df_trn[qdate_col].astype(str).str.contains("2026", na=False))
    df_trn_2026 = df_trn[mask_2026].copy()
    deleted_rows_count = len(df_trn) - len(df_trn_2026)
    report_fn(f"Filter tahun 2026: {len(df_trn_2026):,} baris valid (dihapus {deleted_rows_count:,} baris non-2026)")

    df_trn_2026[qty_col] = to_num(df_trn_2026[qty_col])
    df_trn_2026['item_clean'] = df_trn_2026[it_col].apply(normalize_item)

    trn_grouped = df_trn_2026.groupby('item_clean')[qty_col].sum()
    trn_pivot = {k: v for k, v in trn_grouped.to_dict().items() if k}

    wb_trn = openpyxl.load_workbook(trn_path)
    ws_trn1 = wb_trn["Sheet1"]
    hdr = [str(c.value).strip().lower() if c.value is not None else "" for c in ws_trn1[1]]
    qcol_idx = next(i for i, h in enumerate(hdr, 1) if "quarantine date" in h)
    rows_to_delete = []
    for r_idx in range(2, ws_trn1.max_row + 1):
        v = ws_trn1.cell(row=r_idx, column=qcol_idx).value
        if v is None:
            if all(c.value is None for c in ws_trn1[r_idx]):
                continue
            rows_to_delete.append(r_idx)
            continue
        yr = v.year if hasattr(v, "year") else (2026 if "2026" in str(v) else None)
        if yr != 2026:
            rows_to_delete.append(r_idx)

    blocks = []
    for r_idx in sorted(rows_to_delete):
        if blocks and r_idx == blocks[-1][0] + blocks[-1][1]:
            blocks[-1][1] += 1
        else:
            blocks.append([r_idx, 1])
    for start, length in reversed(blocks):
        ws_trn1.delete_rows(start, length)

    write_pivot_sheet(wb_trn, "Sum of Quantity", trn_pivot)
    safe_save_workbook(wb_trn, trn_path)
    try:
        from native_pivot import create_native_excel_pivot
        create_native_excel_pivot(
            xlsx_path=trn_path,
            table_name="MyGrid",
            target_sheet="Sheet2",
            row_field="Item number",
            val_field="Quantity",
            style_name="PivotStyleLight16",
            log_func=report_fn
        )
    except Exception as e:
        report_fn(f"  [PIVOT NOTE] Menggunakan tabel statis cadangan: {e}")
    report_fn(f"Sheet2 Pivot TRN diperbarui ({len(trn_pivot):,} item unik)")
    return TrnBundle(pivot=trn_pivot)

@dataclass
class SooutBundle:
    pivot_positive: Dict[str, float]

def stage_4_process_soout(soout_path: Optional[str], report_fn: Callable[[str], None]) -> SooutBundle:
    if not soout_path or not os.path.exists(soout_path):
        report_fn("File SOOUT tidak ditemukan.")
        return SooutBundle(pivot_positive={})

    report_fn(f"Membaca file SOOUT: {os.path.basename(soout_path)}...")
    df_soout = pd.read_excel(soout_path, sheet_name="Sheet1", dtype=str)
    it_col = [c for c in df_soout.columns if "item number" in c.lower()][0]
    rem_col = [c for c in df_soout.columns if "deliver reminder" in c.lower() or "deliver remainder" in c.lower()][0]

    df_soout[rem_col] = to_num(df_soout[rem_col])
    df_soout['item_clean'] = df_soout[it_col].apply(normalize_item)

    so_grouped = df_soout.groupby('item_clean')[rem_col].sum()
    so_raw = {k: v for k, v in so_grouped.to_dict().items() if k}
    soout_positive = {k: abs(v) for k, v in so_raw.items()}

    wb_so = openpyxl.load_workbook(soout_path)
    write_pivot_sheet(wb_so, "Sum of Deliver Reminder", so_raw)
    safe_save_workbook(wb_so, soout_path)
    try:
        from native_pivot import create_native_excel_pivot
        create_native_excel_pivot(
            xlsx_path=soout_path,
            table_name="MyGrid",
            target_sheet="Sheet2",
            row_field="Item number",
            val_field="Deliver reminder",
            style_name="PivotStyleLight16",
            log_func=report_fn
        )
    except Exception:
        pass
    report_fn(f"Sheet2 Pivot SOOUT diperbarui ({len(so_raw):,} item unik)")
    return SooutBundle(pivot_positive=soout_positive)

@dataclass
class PooutBundle:
    pivot: Dict[str, float]

def stage_5_process_poout(poout_path: Optional[str], report_fn: Callable[[str], None]) -> PooutBundle:
    if not poout_path or not os.path.exists(poout_path):
        report_fn("File POOUT tidak ditemukan.")
        return PooutBundle(pivot={})

    report_fn(f"Membaca file POOUT: {os.path.basename(poout_path)}...")
    df_po = pd.read_excel(poout_path, sheet_name="Sheet1", dtype=str)
    it_col = [c for c in df_po.columns if "item number" in c.lower()][0]
    rem_col = [c for c in df_po.columns if "deliver remainder" in c.lower() or "deliver reminder" in c.lower()][0]

    df_po[rem_col] = to_num(df_po[rem_col])
    df_po['item_clean'] = df_po[it_col].apply(normalize_item)

    po_grouped = df_po.groupby('item_clean')[rem_col].sum()
    poout_pivot = {k: v for k, v in po_grouped.to_dict().items() if k}

    wb_po = openpyxl.load_workbook(poout_path)
    write_pivot_sheet(wb_po, "Sum of Deliver remainder", poout_pivot)
    safe_save_workbook(wb_po, poout_path)
    try:
        from native_pivot import create_native_excel_pivot
        create_native_excel_pivot(
            xlsx_path=poout_path,
            table_name="MyGrid",
            target_sheet="Sheet2",
            row_field="Item number",
            val_field="Deliver remainder",
            style_name="PivotStyleLight16",
            log_func=report_fn
        )
    except Exception:
        pass
    report_fn(f"Sheet2 Pivot POOUT diperbarui ({len(poout_pivot):,} item unik)")
    return PooutBundle(pivot=poout_pivot)

def stage_6_fill_format_cabang(
    format_cabang_path: str,
    stok: StokBundle,
    trn: TrnBundle,
    soout: SooutBundle,
    poout: PooutBundle,
    report_fn: Callable[[str], None]
) -> int:
    report_fn(f"Membuka Format Cabang: {os.path.basename(format_cabang_path)}...")
    wb_fc = openpyxl.load_workbook(format_cabang_path)
    ws_fc = wb_fc.active

    # Deteksi posisi kolom target secara dinamis dari header baris 3
    col_map = {}
    for c in range(1, ws_fc.max_column + 1):
        v = str(ws_fc.cell(row=3, column=c).value or "").strip().upper()
        if "KODBUK" in v: col_map["KODBUK"] = c
        elif "STOK_PST" in v or "STOK PST" in v: col_map["STOK_PST"] = c
        elif "STOK_NAS" in v or "STOK NAS" in v: col_map["STOK_NAS"] = c
        elif "STOK_MKS" in v or "STOK MKS" in v: col_map["STOK_MKS"] = c
        elif v == "TRANS" or "TRANS" in v: col_map["TRANS"] = c
        elif "SOOUT" in v or "SO OUT" in v: col_map["SOOUT"] = c
        elif "PO OUT" in v or "POOUT" in v: col_map["POOUT"] = c

    c_pst = col_map.get("STOK_PST", 33)
    c_nas = col_map.get("STOK_NAS", 34)
    c_mks = col_map.get("STOK_MKS", 35)
    c_trn = col_map.get("TRANS", 36)
    c_so  = col_map.get("SOOUT", 37)
    c_po  = col_map.get("POOUT", 38)

    consecutive_empty = 0
    actual_last_row = 3
    for r in range(4, min(ws_fc.max_row + 1, 100000)):
        val = ws_fc.cell(row=r, column=1).value
        if val is not None and str(val).strip():
            actual_last_row = r
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            if consecutive_empty > 50:
                break
    last_row = actual_last_row
    total_books = last_row - 3
    report_fn(f"Mengisi {total_books:,} baris buku di Format Cabang (Baris 4 s/d {last_row})...")

    has_pst = bool(stok.pst)
    has_nas = bool(stok.nas)
    has_mks = bool(stok.mks)
    has_trn = bool(trn.pivot)
    has_soout = bool(soout.pivot_positive)
    has_poout = bool(poout.pivot)

    if not has_nas:
        report_fn(f"  [PERHATIAN] Data STOK Nasional tidak tersedia pada file STOK (hanya Pusat & Makassar). Kolom STOK_NAS ({c_nas}) dipertahankan.")
    if not has_pst:
        report_fn(f"  [INFO] Data STOK Pusat tidak tersedia pada file STOK -> Kolom STOK_PST ({c_pst}) dipertahankan.")
    if not has_mks:
        report_fn(f"  [INFO] Data STOK Makassar tidak tersedia pada file STOK -> Kolom STOK_MKS ({c_mks}) dipertahankan.")
    if not has_trn:
        report_fn(f"  [INFO] Data TRN tidak tersedia -> Kolom TRANS ({c_trn}) dipertahankan.")
    if not has_soout:
        report_fn(f"  [INFO] Data SOOUT tidak tersedia -> Kolom SOOUT ({c_so}) dipertahankan.")
    if not has_poout:
        report_fn(f"  [INFO] Data POOUT tidak tersedia -> Kolom PO OUT ({c_po}) dipertahankan.")

    seen_items = set()
    dup_items_found = set()
    for r in range(4, last_row + 1):
        item_val = ws_fc.cell(row=r, column=1).value
        item = normalize_item(item_val)
        if not item:
            continue

        if item in seen_items:
            dup_items_found.add(item)
            # Baris duplikat di Format Cabang: isi 0 untuk mencegah double-count stok fisik gudang
            if has_pst: ws_fc.cell(row=r, column=c_pst, value=0)
            if has_nas: ws_fc.cell(row=r, column=c_nas, value=0)
            if has_mks: ws_fc.cell(row=r, column=c_mks, value=0)
            if has_trn: ws_fc.cell(row=r, column=c_trn, value=0)
            if has_soout: ws_fc.cell(row=r, column=c_so, value=0)
            if has_poout: ws_fc.cell(row=r, column=c_po, value=0)
            continue

        seen_items.add(item)
        if has_pst:
            ws_fc.cell(row=r, column=c_pst, value=int(round(stok.pst.get(item, 0))))
        if has_nas:
            ws_fc.cell(row=r, column=c_nas, value=int(round(stok.nas.get(item, 0))))
        if has_mks:
            ws_fc.cell(row=r, column=c_mks, value=int(round(stok.mks.get(item, 0))))
        if has_trn:
            ws_fc.cell(row=r, column=c_trn, value=int(round(trn.pivot.get(item, 0))))
        if has_soout:
            ws_fc.cell(row=r, column=c_so, value=int(round(soout.pivot_positive.get(item, 0))))
        if has_poout:
            ws_fc.cell(row=r, column=c_po, value=int(round(poout.pivot.get(item, 0))))

    if dup_items_found:
        report_fn(f"  [PROTEKSI DUPLIKASI] Ditemukan {len(dup_items_found)} item duplikat di Format Cabang: {', '.join(sorted(dup_items_found))}. Baris duplikat diisi 0 agar total stok cabang tidak berlebih.")

    safe_save_workbook(wb_fc, format_cabang_path)
    report_fn(f"File Format Cabang berhasil disimpan ({total_books:,} buku terisi)")
    return total_books

def stage_7_process_buku_masuk_fifo(
    inputs: InputBundle,
    report_fn: Callable[[str], None]
) -> Optional[str]:
    bm_paths = inputs.buku_masuk_paths or ([inputs.buku_masuk_path] if inputs.buku_masuk_path else [])
    if not bm_paths:
        report_fn("Tidak ada file Buku Masuk hari ini. Tahap alokasi FIFO dilewati.")
        return None

    if not inputs.soout_path or not os.path.exists(inputs.soout_path):
        raise DailyDataError(
            "File SOOUT tidak ditemukan untuk alokasi Buku Masuk.",
            "Pastikan file SOOUT hari ini ada di folder sebelum memproses Buku Masuk."
        )

    report_fn("Memproses Alokasi Buku Masuk & Ekstraksi FIFO...")
    su_db = {}
    if inputs.su_path and os.path.exists(inputs.su_path):
        su_db = build_sales_hierarchy(inputs.su_path)
        report_fn(f"Master Sales dimuat: {len(su_db):,} sales responsible")

    bm_items = set()
    for bm_p in bm_paths:
        report_fn(f"Membaca file Buku Masuk: {os.path.basename(bm_p)}...")
        df_bm = pd.read_excel(bm_p, dtype=str)
        bm_it_cols = [c for c in df_bm.columns if "item number" in c.lower()]
        if bm_it_cols:
            f_items = {normalize_item(x) for x in df_bm[bm_it_cols[0]].dropna().unique() if normalize_item(x)}
            report_fn(f"  -> {len(f_items):,} item unik terdeteksi pada {os.path.basename(bm_p)}")
            bm_items.update(f_items)
        else:
            report_fn(f"  [PERINGATAN] Kolom 'Item number' tidak ditemukan di {os.path.basename(bm_p)}")

    report_fn(f"Total akumulasi buku masuk: {len(bm_items):,} item unik dari {len(bm_paths)} file.")

    df_so_raw = pd.read_excel(inputs.soout_path, sheet_name="Sheet1", dtype=str)
    so_col = [c for c in df_so_raw.columns if "sales order" in c.lower()][0]
    date_col = [c for c in df_so_raw.columns if "created date" in c.lower()][0]
    it_col = [c for c in df_so_raw.columns if "item number" in c.lower()][0]
    resp_col = [c for c in df_so_raw.columns if "sales responsible" in c.lower()][0]

    df_soe = df_so_raw[df_so_raw[so_col].astype(str).str.startswith("56-SOE")].copy()
    df_soe['item_clean'] = df_soe[it_col].apply(normalize_item)
    df_soe['dt_sort'] = pd.to_datetime(df_soe[date_col], format='mixed', dayfirst=True, errors='coerce')
    df_soe = df_soe.sort_values(by=['item_clean', 'dt_sort'], ascending=[True, True])

    gl_list, asm_list, mgr_list = [], [], []
    for _, row in df_soe.iterrows():
        resp = str(row[resp_col]).strip()
        lookup = su_db.get(resp, {"GL": "", "ASM": "", "MGR": ""})
        gl_list.append(lookup["GL"])
        asm_list.append(lookup["ASM"])
        mgr_list.append(lookup["MGR"])

    col_idx = 5
    df_soe.insert(col_idx, "GL", gl_list)
    df_soe.insert(col_idx + 1, "ASM", asm_list)
    df_soe.insert(col_idx + 2, "MGR", mgr_list)

    wb_so = openpyxl.load_workbook(inputs.soout_path)
    if "Sheet3" in wb_so.sheetnames:
        ws_so3 = wb_so["Sheet3"]
        ws_so3.delete_rows(1, ws_so3.max_row + 1)
    else:
        ws_so3 = wb_so.create_sheet("Sheet3")

    df_soe_save = df_soe.drop(columns=['item_clean', 'dt_sort'])
    for r_idx, row in enumerate(dataframe_to_rows(df_soe_save, index=False, header=True), 1):
        for c_idx, val in enumerate(row, 1):
            ws_so3.cell(row=r_idx, column=c_idx, value=val)
    safe_save_workbook(wb_so, inputs.soout_path)
    report_fn(f"Sheet3 SOOUT FIFO disimpan ({len(df_soe_save):,} baris SOE)")

    # Cek apakah file print sudah pernah dibuat sebelumnya
    bm_print_candidates = [
        f for f in glob.glob(os.path.join(inputs.work_dir, "BUKMAS*PRINT*.xlsx"))
        if "(BARU)" not in os.path.basename(f).upper()
    ]
    bm_print_candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
    existing_print_path = bm_print_candidates[0] if bm_print_candidates else None
    old_items_list = []
    existing_items = set()
    existing_sos = set()

    if existing_print_path and os.path.exists(existing_print_path):
        try:
            wb_prev = openpyxl.load_workbook(existing_print_path, data_only=True, read_only=True)
            if "ITEM_MASUK" in wb_prev.sheetnames:
                ws_it = wb_prev["ITEM_MASUK"]
                for r_idx, row in enumerate(ws_it.iter_rows(values_only=True), 1):
                    if r_idx > 1 and row and row[0]:
                        v_str = str(row[0]).strip()
                        if "new" in v_str.lower() and "=" in v_str:
                            continue
                        clean_it = normalize_item(row[0])
                        if clean_it:
                            existing_items.add(clean_it)
                            if clean_it not in old_items_list:
                                old_items_list.append(clean_it)
            for sname in ["MGR1", "MGR2"]:
                if sname in wb_prev.sheetnames:
                    ws_prev = wb_prev[sname]
                    it_idx, so_idx = None, None
                    for r_idx, r_vals in enumerate(ws_prev.iter_rows(values_only=True), 1):
                        if r_idx == 1:
                            for c_i, h in enumerate(r_vals):
                                h_str = str(h or "").lower()
                                if "item number" in h_str and it_idx is None:
                                    it_idx = c_i
                                if "sales order" in h_str and so_idx is None:
                                    so_idx = c_i
                            continue
                        if not r_vals:
                            continue
                        if it_idx is not None and it_idx < len(r_vals) and r_vals[it_idx]:
                            existing_items.add(normalize_item(r_vals[it_idx]))
                        if so_idx is not None and so_idx < len(r_vals) and r_vals[so_idx]:
                            existing_sos.add(str(r_vals[so_idx]).strip())
            wb_prev.close()
        except Exception as e:
            report_fn(f"  [INFO] Catatan pembacaan file print lama: {e}")

    df_filtered = df_soe[df_soe['item_clean'].isin(bm_items)].copy()
    current_sos = set(df_filtered[so_col].astype(str).str.strip().unique()) if not df_filtered.empty else set()

    # Evaluasi data baru
    new_bm_items = bm_items - existing_items
    new_sos = current_sos - existing_sos

    # Cek apakah item baru memiliki pesanan SO terbuka
    new_items_with_so = [it for it in new_bm_items if it in df_soe['item_clean'].values]

    # Jika file print sudah ada dan TIDAK ADA buku masuk baru maupun SO baru, lewati ekspor
    if existing_print_path and len(existing_items) > 0 and len(new_bm_items) == 0 and len(new_sos) == 0:
        report_fn(f"  [INFO] Tidak ada buku masuk baru yang belum diekspor.")
        report_fn(f"  [INFO] Seluruh {len(bm_items):,} item buku masuk ({len(current_sos):,} SO) sudah tercakup dalam {os.path.basename(existing_print_path)}.")
        report_fn(f"  [INFO] Ekspor dilewati karena semua data sudah mutakhir (aman dari overwrite).")
        return existing_print_path

    is_incremental = bool(existing_print_path and len(existing_items) > 0 and (len(new_bm_items) > 0 or len(new_sos) > 0))
    if is_incremental:
        report_fn(f"  [UPDATE] Terdeteksi kiriman buku masuk baru ({len(new_items_with_so):,} item baru ber-SO, {len(new_sos):,} pesanan SO baru)!")
        if new_items_with_so:
            report_fn(f"  [UPDATE] Contoh item baru: {', '.join(sorted(list(new_items_with_so))[:5])}")

    df_m1 = df_filtered[df_filtered['MGR'].str.contains("1", na=False)].copy()
    df_m2 = df_filtered[df_filtered['MGR'].str.contains("2", na=False)].copy()

    cols_28 = list(df_soe_save.columns[:28])

    if is_incremental:
        df_filtered_old = df_filtered[df_filtered['item_clean'].isin(existing_items) & ~df_filtered[so_col].isin(new_sos)].copy()
        df_filtered_new = df_filtered[df_filtered['item_clean'].isin(new_bm_items) | df_filtered[so_col].isin(new_sos)].copy()

        df_m1_old = df_filtered_old[df_filtered_old['MGR'].str.contains("1", na=False)][cols_28].copy()
        df_m1_new = df_filtered_new[df_filtered_new['MGR'].str.contains("1", na=False)][cols_28].copy()

        df_m2_old = df_filtered_old[df_filtered_old['MGR'].str.contains("2", na=False)][cols_28].copy()
        df_m2_new = df_filtered_new[df_filtered_new['MGR'].str.contains("2", na=False)][cols_28].copy()

        so_m1_old_list = [s for s in dict.fromkeys(df_m1_old[so_col])]
        so_m1_new_list = [s for s in dict.fromkeys(df_m1_new[so_col])]

        so_m2_old_list = [s for s in dict.fromkeys(df_m2_old[so_col])]
        so_m2_new_list = [s for s in dict.fromkeys(df_m2_new[so_col])]
    else:
        df_m1 = df_filtered[df_filtered['MGR'].str.contains("1", na=False)][cols_28].copy()
        df_m2 = df_filtered[df_filtered['MGR'].str.contains("2", na=False)][cols_28].copy()

        df_m1_old = df_m1
        df_m1_new = pd.DataFrame(columns=cols_28)

        df_m2_old = df_m2
        df_m2_new = pd.DataFrame(columns=cols_28)

        so_m1_old_list = [s for s in dict.fromkeys(df_m1[so_col])]
        so_m1_new_list = []

        so_m2_old_list = [s for s in dict.fromkeys(df_m2[so_col])]
        so_m2_new_list = []

    bukmas_filename = f"BUKMAS {inputs.date_str} PRINT.xlsx"
    bukmas_print_path = os.path.join(inputs.work_dir, bukmas_filename)
    wb_bp = openpyxl.Workbook()

    # Sheet MGR1: Data lama lalu data baru langsung di bawahnya tanpa pemisah
    ws_m1 = wb_bp.active
    ws_m1.title = "MGR1"
    for c_idx, col_name in enumerate(cols_28, 1):
        ws_m1.cell(row=1, column=c_idx, value=col_name)
    cur_m1_row = 2
    for row in dataframe_to_rows(df_m1_old, index=False, header=False):
        for c_idx, val in enumerate(row, 1):
            ws_m1.cell(row=cur_m1_row, column=c_idx, value=val)
        cur_m1_row += 1
    if is_incremental and not df_m1_new.empty:
        for row in dataframe_to_rows(df_m1_new, index=False, header=False):
            for c_idx, val in enumerate(row, 1):
                ws_m1.cell(row=cur_m1_row, column=c_idx, value=val)
            cur_m1_row += 1
    ws_m1.auto_filter.ref = ws_m1.dimensions

    # Sheet MGR2: Data lama lalu data baru langsung di bawahnya tanpa pemisah
    ws_m2 = wb_bp.create_sheet("MGR2")
    for c_idx, col_name in enumerate(cols_28, 1):
        ws_m2.cell(row=1, column=c_idx, value=col_name)
    cur_m2_row = 2
    for row in dataframe_to_rows(df_m2_old, index=False, header=False):
        for c_idx, val in enumerate(row, 1):
            ws_m2.cell(row=cur_m2_row, column=c_idx, value=val)
        cur_m2_row += 1
    if is_incremental and not df_m2_new.empty:
        for row in dataframe_to_rows(df_m2_new, index=False, header=False):
            for c_idx, val in enumerate(row, 1):
                ws_m2.cell(row=cur_m2_row, column=c_idx, value=val)
            cur_m2_row += 1
    ws_m2.auto_filter.ref = ws_m2.dimensions

    # Sheet DELIMIT: Penanda simpel =====new===== hanya jika ada data baru
    ws_delimit = wb_bp.create_sheet("DELIMIT")
    ws_delimit.cell(row=1, column=1, value="PANDUAN COPY-PASTE FILTER DYNAMICS AX (PER 14 ITEM)")
    ws_delimit.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True, size=12)

    cur_row = 3
    # M1 Delimit (Awal / Lama)
    ws_delimit.cell(row=cur_row, column=1, value=f"[ MANAGER 1 - TOTAL {len(so_m1_old_list)} SALES ORDER ]").font = openpyxl.styles.Font(bold=True, color="000080")
    cur_row += 1
    batch_num = 1
    for i in range(0, len(so_m1_old_list), 14):
        chunk = so_m1_old_list[i:i+14]
        ws_delimit.cell(row=cur_row, column=1, value=f"Batch M1-{batch_num} ({len(chunk)} item):")
        ws_delimit.cell(row=cur_row, column=2, value=",".join(chunk))
        cur_row += 1
        batch_num += 1

    cur_row += 1
    # M2 Delimit (Awal / Lama)
    ws_delimit.cell(row=cur_row, column=1, value=f"[ MANAGER 2 - TOTAL {len(so_m2_old_list)} SALES ORDER ]").font = openpyxl.styles.Font(bold=True, color="000080")
    cur_row += 1
    batch_num = 1
    for i in range(0, len(so_m2_old_list), 14):
        chunk = so_m2_old_list[i:i+14]
        ws_delimit.cell(row=cur_row, column=1, value=f"Batch M2-{batch_num} ({len(chunk)} item):")
        ws_delimit.cell(row=cur_row, column=2, value=",".join(chunk))
        cur_row += 1
        batch_num += 1

    # Jika incremental dan ada SO baru, beri penanda simpel =====new=====
    if is_incremental and (so_m1_new_list or so_m2_new_list):
        cur_row += 1
        ws_delimit.cell(row=cur_row, column=1, value="=====new=====").font = openpyxl.styles.Font(bold=True, color="B91C1C", size=11)
        cur_row += 2

        if so_m1_new_list:
            ws_delimit.cell(row=cur_row, column=1, value=f"[ MANAGER 1 - TOTAL {len(so_m1_new_list)} SALES ORDER ]").font = openpyxl.styles.Font(bold=True, color="000080")
            cur_row += 1
            for i in range(0, len(so_m1_new_list), 14):
                chunk = so_m1_new_list[i:i+14]
                ws_delimit.cell(row=cur_row, column=1, value=f"Batch M1-{(i//14)+1} ({len(chunk)} item):")
                ws_delimit.cell(row=cur_row, column=2, value=",".join(chunk))
                cur_row += 1

        if so_m2_new_list:
            if so_m1_new_list:
                cur_row += 1
            ws_delimit.cell(row=cur_row, column=1, value=f"[ MANAGER 2 - TOTAL {len(so_m2_new_list)} SALES ORDER ]").font = openpyxl.styles.Font(bold=True, color="000080")
            cur_row += 1
            for i in range(0, len(so_m2_new_list), 14):
                chunk = so_m2_new_list[i:i+14]
                ws_delimit.cell(row=cur_row, column=1, value=f"Batch M2-{(i//14)+1} ({len(chunk)} item):")
                ws_delimit.cell(row=cur_row, column=2, value=",".join(chunk))
                cur_row += 1

    # Sheet ITEM_MASUK: Penanda simpel =====new===== hanya jika ada data baru
    ws_it_save = wb_bp.create_sheet("ITEM_MASUK")
    ws_it_save.cell(row=1, column=1, value="Item number")
    cur_it_row = 2

    if is_incremental:
        items_to_write_old = old_items_list if old_items_list else [it for it in sorted(existing_items)]
        if not items_to_write_old:
            items_to_write_old = [it for it in sorted(bm_items - set(new_bm_items))]
        for it_code in items_to_write_old:
            ws_it_save.cell(row=cur_it_row, column=1, value=it_code)
            cur_it_row += 1

        if new_bm_items:
            ws_it_save.cell(row=cur_it_row, column=1, value="=====new=====").font = openpyxl.styles.Font(bold=True, color="B91C1C")
            cur_it_row += 1
            for it_code in sorted(new_bm_items):
                ws_it_save.cell(row=cur_it_row, column=1, value=it_code)
                cur_it_row += 1
    else:
        for it_code in sorted(bm_items):
            ws_it_save.cell(row=cur_it_row, column=1, value=it_code)
            cur_it_row += 1

    for ws in wb_bp.worksheets:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col[:40])
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 45)

    safe_save_workbook(wb_bp, bukmas_print_path)
    if is_incremental:
        report_fn(f"  [SUKSES] Master print diperbarui: {bukmas_filename} (+{len(df_m1_new)} M1, +{len(df_m2_new)} M2 baru, penanda '=====new=====')")
    else:
        report_fn(f"  [SUKSES] File Print selesai dibuat: {bukmas_filename} (M1: {len(df_m1_old)} baris, M2: {len(df_m2_old)} baris)")

    return bukmas_print_path

def stage_8_process_do_summary(
    do_summary_path: Optional[str],
    work_dir: str,
    report_fn: Callable[[str], None],
    date_str: Optional[str] = None
) -> Optional[str]:
    if not do_summary_path or not os.path.exists(do_summary_path):
        do_summary_path = find_do_summary_for_date(work_dir, target_date_str=date_str, report_fn=report_fn)
    if not do_summary_path or not os.path.exists(do_summary_path):
        report_fn(f"[SKIP] File DO Summary untuk tanggal {date_str or ''} belum ditemukan di folder kerja maupun Downloads.")
        return None
    out_xlsx = os.path.join(work_dir, os.path.basename(do_summary_path) + ".xlsx")
    report_fn(f"Memproses DO Summary: {os.path.basename(do_summary_path)}...")
    return process_do_summary_file(do_summary_path, out_xlsx_path=out_xlsx, log_func=report_fn)

def stage_9_process_recheck_sp(work_dir: str, report_fn: Callable[[str], None]) -> Optional[str]:
    report_fn("Memeriksa kesesuaian SP kemarin & mendeteksi potensi selisih stok...")
    res = run_recheck_sp(work_dir, progress_fn=report_fn)
    return res.excel_output_path

def stage_10_sync_examples(files_to_sync: List[Optional[str]], report_fn: Callable[[str], None]):
    os.makedirs(EXAMPLES_DIR, exist_ok=True)
    report_fn(f"Menyinkronkan output ke folder: {os.path.relpath(EXAMPLES_DIR, PROJECT_ROOT)}")
    for src in files_to_sync:
        if src and os.path.exists(src):
            dest_name = os.path.basename(src)
            dest = os.path.join(EXAMPLES_DIR, dest_name)
            try:
                shutil.copy2(src, dest)
                repair_external_links(dest)
                report_fn(f"  - Tersinkron: {dest_name}")
            except Exception as e:
                report_fn(f"  - Catatan sinkronisasi {dest_name}: {e}")

# ==============================================================================
# 5. PIPELINE CONTROLLER & ORCHESTRATOR
# ==============================================================================
class DailyDataPipeline:
    """
    Enterprise Orchestrator:
    Mengelola urutan timeline, isolasi tahapan, pelaporan real-time,
    dan diagnosis kegagalan berstandar industri.
    """
    def __init__(
        self,
        work_dir: Optional[str] = None,
        mode: str = PipelineMode.ALL,
        selected_stages: Optional[List[int]] = None,
        progress_callback: Optional[Callable[[str], None]] = None,
        stage_callback: Optional[Callable[[StageMetric], None]] = None
    ):
        self.work_dir_input = work_dir
        self.mode = mode
        self.selected_stages = selected_stages
        self.progress_callback = progress_callback
        self.stage_callback = stage_callback
        self.timeline = ExecutionTimeline()
        self._init_stages()

    def _init_stages(self):
        self.idx_init = self.timeline.register_stage("STAGE_01_INIT", "Deteksi & Validasi File Input")
        self.idx_stok = self.timeline.register_stage("STAGE_02_STOK", "Agregasi & Pivot STOK")
        self.idx_trn = self.timeline.register_stage("STAGE_03_TRN", "Filter 2026 Murni & Pivot TRN")
        self.idx_so = self.timeline.register_stage("STAGE_04_SOOUT", "Pivot SOOUT (Deliver Reminder)")
        self.idx_po = self.timeline.register_stage("STAGE_05_POOUT", "Pivot POOUT (Deliver Remainder)")
        self.idx_fc = self.timeline.register_stage("STAGE_06_FORMAT_CABANG", "Pengisian Nilai Format Cabang")
        self.idx_bm = self.timeline.register_stage("STAGE_07_BUKU_MASUK_FIFO", "Alokasi Buku Masuk & Ekstraksi FIFO")
        self.idx_do = self.timeline.register_stage("STAGE_08_DO_SUMMARY", "Pivot Tabular DO Summary External")
        self.idx_recheck = self.timeline.register_stage("STAGE_09_RECHECK_SP", "Recheck SP Kemarin & Validasi Posting")

    def report(self, message: str, level: str = "INFO"):
        prefix = f"[{level}] " if level != "INFO" else ""
        formatted = f"{prefix}{message}"
        log(formatted)
        if self.progress_callback:
            self.progress_callback(formatted)

    def _execute_stage(self, stage_idx: int, func: Callable, *args, **kwargs) -> Any:
        stage = self.timeline.get_stage(stage_idx)
        total_stages = len(self.timeline.stages)
        self.timeline.start_stage(stage_idx)
        self.report(f"\n==================================================", "STAGE")
        self.report(f"TAHAP [{stage.index}/{total_stages}]: {stage.name.upper()}", "STAGE")
        self.report(f"==================================================", "STAGE")
        if self.stage_callback:
            self.stage_callback(stage)

        try:
            result = func(*args, **kwargs)
            summary = ""
            if isinstance(result, (int, float, str)):
                summary = str(result)
            elif isinstance(result, InputBundle):
                summary = f"Tanggal {result.date_str}"
            elif isinstance(result, StokBundle):
                summary = f"{len(result.pst):,} PST | {len(result.nas):,} NAS | {len(result.mks):,} MKS"
            elif isinstance(result, TrnBundle):
                summary = f"{len(result.pivot):,} item 2026"
            elif isinstance(result, SooutBundle):
                summary = f"{len(result.pivot_positive):,} item SO"
            elif isinstance(result, PooutBundle):
                summary = f"{len(result.pivot):,} item PO"

            self.timeline.complete_stage(stage_idx, summary)
            self.report(f"[OK] Tahap [{stage.index}/{total_stages}] Selesai ({stage.duration_sec:.2f}s)", "SUCCESS")
            if self.stage_callback:
                self.stage_callback(stage)
            return result
        except Exception as e:
            tb = traceback.format_exc()
            self.timeline.fail_stage(stage_idx, e, tb)
            if self.stage_callback:
                self.stage_callback(stage)

            err_banner = [
                "",
                "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!",
                f"[ERROR PIPELINE] KEGAGALAN PADA TAHAP [{stage.index}/9]: {stage.name}",
                "------------------------------------------------------------------------",
                f"Waktu Kejadian : {datetime.now().strftime('%H:%M:%S')} (Durasi berjalan: {stage.duration_sec:.2f}s)",
                f"Tipe Error     : {type(e).__name__}",
                f"Pesan Error    : {str(e)}",
            ]
            tip = getattr(e, "resolution_tip", "")
            if tip:
                err_banner.append(f"Saran Solusi   : {tip}")
            err_banner.append("------------------------------------------------------------------------")
            err_banner.append(f"Detail Traceback:\n{tb.strip()}")
            err_banner.append("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
            err_banner_str = "\n".join(err_banner)

            self.report(err_banner_str, "ERROR")
            raise StageExecutionError(stage.index, stage.name, e, tb) from e

    def run(self) -> ExecutionTimeline:
        pipeline_start = time.perf_counter()

        # 1. Tentukan target stages
        if self.selected_stages:
            target_stages = set(self.selected_stages)
            if 1 not in target_stages:
                target_stages.add(1)
        elif self.mode in MODE_DEFAULT_STAGES:
            target_stages = set(MODE_DEFAULT_STAGES[self.mode])
        else:
            target_stages = set(MODE_DEFAULT_STAGES[PipelineMode.ALL])

        mode_name = MODE_DESCRIPTIONS.get(self.mode, str(self.mode))
        self.report("================================================================================")
        self.report("                     MEMULAI PIPELINE HARIAN: DAILY DATA                        ")
        self.report("================================================================================")
        self.report(f"Mode Operasi    : {mode_name}")
        target_str = ", ".join([f"Tahap {s}" for s in sorted(target_stages)])
        self.report(f"Tahapan Aktif   : {target_str}")

        # 2. Resolusi Folder Kerja
        if self.work_dir_input:
            work_dir = resolve_work_dir(self.work_dir_input)
        else:
            work_dir = os.getcwd()
            if not glob.glob(os.path.join(work_dir, "FORMAT POSISISTOK*.xlsx")):
                dated_dirs = list_date_folders(PROJECT_ROOT)
                if dated_dirs:
                    work_dir = dated_dirs[0]

        self.report(f"Direktori Target: {work_dir}")

        # Tahap 1: Deteksi File Input (Selalu jalan untuk discovery)
        inputs: InputBundle = self._execute_stage(self.idx_init, stage_1_discover_inputs, work_dir, self.report, sorted(target_stages))

        # Tahap 2: Pivot STOK
        if 2 in target_stages:
            stok_res: StokBundle = self._execute_stage(self.idx_stok, stage_2_process_stok, inputs.stok_path, self.report)
        else:
            self.timeline.skip_stage(self.idx_stok, "Dilewati (Tidak dipilih)")
            stok_res = StokBundle(pst={}, nas={}, mks={})

        # Tahap 3: Pivot TRN
        if 3 in target_stages:
            trn_res: TrnBundle = self._execute_stage(self.idx_trn, stage_3_process_trn, inputs.trn_path, self.report)
        else:
            self.timeline.skip_stage(self.idx_trn, "Dilewati (Tidak dipilih)")
            trn_res = TrnBundle(pivot={})

        # Tahap 4: Pivot SOOUT
        if 4 in target_stages:
            soout_res: SooutBundle = self._execute_stage(self.idx_so, stage_4_process_soout, inputs.soout_path, self.report)
        else:
            self.timeline.skip_stage(self.idx_so, "Dilewati (Tidak dipilih)")
            soout_res = SooutBundle(pivot_positive={})

        # Tahap 5: Pivot POOUT
        if 5 in target_stages:
            poout_res: PooutBundle = self._execute_stage(self.idx_po, stage_5_process_poout, inputs.poout_path, self.report)
        else:
            self.timeline.skip_stage(self.idx_po, "Dilewati (Tidak dipilih)")
            poout_res = PooutBundle(pivot={})

        # Tahap 6: Format Cabang
        if 6 in target_stages:
            if not inputs.format_cabang_path:
                raise MandatoryFileNotFoundError("FORMAT POSISISTOK*.xlsx", work_dir)
            total_books = self._execute_stage(
                self.idx_fc,
                stage_6_fill_format_cabang,
                inputs.format_cabang_path,
                stok_res,
                trn_res,
                soout_res,
                poout_res,
                self.report
            )
        else:
            self.timeline.skip_stage(self.idx_fc, "Dilewati (Format Cabang tidak diubah)")

        # Tahap 7: Buku Masuk & FIFO (Kondisional)
        bukmas_out = None
        if 7 in target_stages:
            if inputs.buku_masuk_paths:
                bukmas_out = self._execute_stage(self.idx_bm, stage_7_process_buku_masuk_fifo, inputs, self.report)
            else:
                self.timeline.skip_stage(self.idx_bm, "Tidak ada file Buku Masuk hari ini")
                self.report(f"\nTahap [7/9] Alokasi Buku Masuk: Dilewati (Tidak ada penerimaan buku hari ini).", "INFO")
        else:
            self.timeline.skip_stage(self.idx_bm, "Dilewati (Tidak dipilih)")

        # Tahap 8: DO Summary External (Kondisional)
        do_out = None
        if 8 in target_stages:
            curr_do = inputs.do_summary_path
            if not curr_do or not os.path.exists(curr_do):
                curr_do = find_do_summary_for_date(inputs.work_dir, target_date_str=inputs.date_str, report_fn=self.report)
                inputs.do_summary_path = curr_do
            if curr_do and os.path.exists(curr_do):
                do_out = self._execute_stage(self.idx_do, stage_8_process_do_summary, curr_do, inputs.work_dir, self.report, inputs.date_str)
            else:
                self.timeline.skip_stage(self.idx_do, "Tidak ada file DO Summary sore ini")
                self.report(f"\nTahap [8/9] DO Summary: Dilewati (File DO Summary untuk tanggal {inputs.date_str} belum diunduh / tidak ditemukan).", "INFO")
        else:
            self.timeline.skip_stage(self.idx_do, "Dilewati (Tidak dipilih)")

        # Tahap 9: Recheck SP Kemarin & Validasi Posting
        recheck_out = None
        if 9 in target_stages:
            recheck_out = self._execute_stage(self.idx_recheck, stage_9_process_recheck_sp, inputs.work_dir, self.report)
        else:
            self.timeline.skip_stage(self.idx_recheck, "Dilewati (Tidak dipilih)")

        # Timeline Summary
        summary_text = self.timeline.get_summary_text()
        self.report("\n" + summary_text)
        return self.timeline

def run_automation(
    work_dir: Optional[str] = None,
    mode: str = PipelineMode.ALL,
    selected_stages: Optional[List[int]] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    stage_callback: Optional[Callable[[StageMetric], None]] = None
) -> ExecutionTimeline:
    pipeline = DailyDataPipeline(
        work_dir=work_dir,
        mode=mode,
        selected_stages=selected_stages,
        progress_callback=progress_callback,
        stage_callback=stage_callback
    )
    return pipeline.run()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Daily Data Automation Pipeline")
    parser.add_argument("folder", nargs="?", default=None, help="Folder target harian (misal: 07 OKT 2026)")
    parser.add_argument("--mode", choices=["all", "stok", "recheck", "bukmas", "dosum", "custom"], default="all",
                        help="Mode alur kerja: all (1-10), stok (1-6, 9-10), recheck (1, 9), bukmas (7), dosum (8)")
    parser.add_argument("--stages", help="Daftar nomor tahap dipisah koma (misal: 1,9,10)")
    parser.add_argument("--inspect", help="Pemeriksaan instan item buku & antrean SO FIFO (Kode atau Judul)")
    parser.add_argument("--recheck", action="store_true", help="Jalankan recheck SP kemarin & validasi posting instan")
    args = parser.parse_args()

    if args.recheck:
        w_dir = resolve_work_dir(args.folder) if args.folder else os.getcwd()
        if not glob.glob(os.path.join(w_dir, "FORMAT POSISISTOK*.xlsx")):
            dated_dirs = list_date_folders(PROJECT_ROOT)
            if dated_dirs:
                w_dir = dated_dirs[0]
        run_recheck_sp(w_dir)
        sys.exit(0)

    if args.inspect:
        from item_inspector import inspect_item
        w_dir = resolve_work_dir(args.folder) if args.folder else os.getcwd()
        if not glob.glob(os.path.join(w_dir, "FORMAT POSISISTOK*.xlsx")):
            dated_dirs = list_date_folders(PROJECT_ROOT)
            if dated_dirs:
                w_dir = dated_dirs[0]
        res = inspect_item(args.inspect, w_dir)
        if res:
            print(res.formatted_report)
        else:
            print(f"[INFO] Item '{args.inspect}' tidak ditemukan di folder: {w_dir}")
        sys.exit(0)

    mode_map = {
        "all": PipelineMode.ALL,
        "stok": PipelineMode.POSISI_STOK,
        "recheck": PipelineMode.RECHECK_SP,
        "bukmas": PipelineMode.BUKU_MASUK,
        "dosum": PipelineMode.DO_SUMMARY,
        "custom": PipelineMode.CUSTOM
    }
    sel_stages = [int(s.strip()) for s in args.stages.split(",") if s.strip().isdigit()] if args.stages else None

    try:
        run_automation(work_dir=args.folder, mode=mode_map.get(args.mode, PipelineMode.ALL), selected_stages=sel_stages)
    except Exception as exc:
        sys.exit(1)
