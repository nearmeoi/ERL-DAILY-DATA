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

from process_do_summary import process_do_summary_file

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

def write_pivot_sheet(wb: openpyxl.Workbook, header_value: str, pivot_dict: Dict[str, float], sheet_name: str = "Sheet2"):
    """
    Menulis hasil pivot tabel standar:
    Diletakkan di tab pertama, baris 3 header, tanpa item kosong, dan Grand Total di akhir.
    """
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name, 0)
    ws.cell(row=3, column=1, value="Row Labels")
    ws.cell(row=3, column=2, value=header_value)
    r = 4
    total = 0.0
    for item, q in sorted((k, v) for k, v in pivot_dict.items() if k):
        ws.cell(row=r, column=1, value=item)
        ws.cell(row=r, column=2, value=q)
        total += q
        r += 1
    ws.cell(row=r, column=1, value="Grand Total")
    ws.cell(row=r, column=2, value=total)
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
    format_cabang_path: str
    stok_path: Optional[str]
    trn_path: Optional[str]
    soout_path: Optional[str]
    poout_path: Optional[str]
    su_path: Optional[str]
    buku_masuk_path: Optional[str]
    do_summary_path: Optional[str]

def stage_1_discover_inputs(work_dir: str, report_fn: Callable[[str], None]) -> InputBundle:
    report_fn(f"Menelaah file input pada folder: {work_dir}")
    format_cabang_path = find_file("FORMAT POSISISTOK*.xlsx", [work_dir])
    if not format_cabang_path:
        raise MandatoryFileNotFoundError("FORMAT POSISISTOK*.xlsx", work_dir)

    trn_path = find_file("TRN*.xlsx", [work_dir, "D:\\ANALISA STOK"])
    soout_path = find_file("SOOUT*.xlsx", [work_dir])
    poout_path = find_file("POOUT*.xlsx", [work_dir])
    stok_path = find_file("STOK*.xlsx", [work_dir, "D:\\ANALISA STOK", "D:\\"])

    ref_dir = os.path.join(PROJECT_ROOT, "REFERENSI")
    su_path = find_file("SALESUNIT2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\", APP_DIR]) or \
              find_file("SU2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\", APP_DIR])

    buku_masuk_path = find_file("buku masuk*.xlsx", [work_dir])
    if not buku_masuk_path:
        candidates = glob.glob(os.path.join(work_dir, "BUKMAS*.xlsx"))
        candidates = [c for c in candidates if "PRINT" not in os.path.basename(c).upper()]
        if candidates:
            candidates.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            buku_masuk_path = candidates[0]

    do_summary_path = find_file("Do_Summary_External*.csv", [work_dir])
    date_str = extract_date_str(format_cabang_path)

    report_fn(f"  - Format Cabang : {os.path.basename(format_cabang_path)}")
    report_fn(f"  - STOK          : {os.path.basename(stok_path) if stok_path else '[TIDAK DITEMUKAN]'}")
    report_fn(f"  - TRN           : {os.path.basename(trn_path) if trn_path else '[TIDAK DITEMUKAN]'}")
    report_fn(f"  - SOOUT         : {os.path.basename(soout_path) if soout_path else '[TIDAK DITEMUKAN]'}")
    report_fn(f"  - POOUT         : {os.path.basename(poout_path) if poout_path else '[TIDAK DITEMUKAN]'}")
    report_fn(f"  - Master Sales  : {os.path.basename(su_path) if su_path else '[TIDAK DITEMUKAN]'}")
    report_fn(f"  - Buku Masuk    : {os.path.basename(buku_masuk_path) if buku_masuk_path else '[TIDAK ADA - SKIP]'}")
    report_fn(f"  - DO Summary    : {os.path.basename(do_summary_path) if do_summary_path else '[TIDAK ADA - SKIP]'}")
    report_fn(f"  - Tanggal Opr   : {date_str}")

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

    df_pst = df_stok[df_stok['wh_clean'] == '50'].groupby('item_clean')[qty_col].sum()
    stok_pst = df_pst.to_dict()

    df_mks = df_stok[df_stok['wh_clean'] == '56'].groupby('item_clean')[qty_col].sum()
    stok_mks = df_mks.to_dict()

    # Rule Baku: Nasional = Semua kecuali 50, 56, 723
    df_nas = df_stok[~df_stok['wh_clean'].isin(['50', '56', '723'])].groupby('item_clean')[qty_col].sum()
    stok_nas = df_nas.to_dict()

    report_fn(f"Agregasi selesai: {len(stok_pst):,} item Pusat, {len(stok_nas):,} item Nasional, {len(stok_mks):,} item Makassar")
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

    for r in range(4, last_row + 1):
        item_val = ws_fc.cell(row=r, column=1).value
        item = normalize_item(item_val)
        if not item:
            continue
        ws_fc.cell(row=r, column=33, value=stok.pst.get(item, 0))
        ws_fc.cell(row=r, column=34, value=stok.nas.get(item, 0))
        ws_fc.cell(row=r, column=35, value=stok.mks.get(item, 0))
        ws_fc.cell(row=r, column=36, value=trn.pivot.get(item, 0))
        ws_fc.cell(row=r, column=37, value=soout.pivot_positive.get(item, 0))
        ws_fc.cell(row=r, column=38, value=poout.pivot.get(item, 0))

    safe_save_workbook(wb_fc, format_cabang_path)
    report_fn(f"File Format Cabang berhasil disimpan ({total_books:,} buku terisi)")
    return total_books

def stage_7_process_buku_masuk_fifo(
    inputs: InputBundle,
    report_fn: Callable[[str], None]
) -> Optional[str]:
    if not inputs.buku_masuk_path or not os.path.exists(inputs.buku_masuk_path):
        report_fn("Tidak ada file Buku Masuk hari ini. Tahap alokasi FIFO dilewati.")
        return None

    if not inputs.soout_path or not os.path.exists(inputs.soout_path):
        raise DailyDataError("File SOOUT tidak ditemukan untuk alokasi Buku Masuk.")

    report_fn("Memproses Alokasi Buku Masuk & Ekstraksi FIFO...")
    su_db = {}
    if inputs.su_path and os.path.exists(inputs.su_path):
        su_db = build_sales_hierarchy(inputs.su_path)
        report_fn(f"Master Sales dimuat: {len(su_db):,} sales responsible")

    df_bm = pd.read_excel(inputs.buku_masuk_path, dtype=str)
    bm_it_col = [c for c in df_bm.columns if "item number" in c.lower()][0]
    bm_items = {normalize_item(x) for x in df_bm[bm_it_col].dropna().unique()}
    report_fn(f"Buku masuk memuat {len(bm_items):,} item unik.")

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

    df_filtered = df_soe[df_soe['item_clean'].isin(bm_items)].copy()
    df_m1 = df_filtered[df_filtered['MGR'].str.contains("1", na=False)].copy()
    df_m2 = df_filtered[df_filtered['MGR'].str.contains("2", na=False)].copy()

    cols_28 = list(df_soe_save.columns[:28])
    df_m1_print = df_m1[cols_28].copy()
    df_m2_print = df_m2[cols_28].copy()

    bukmas_filename = f"BUKMAS {inputs.date_str} PRINT.xlsx"
    bukmas_print_path = os.path.join(inputs.work_dir, bukmas_filename)
    wb_bp = openpyxl.Workbook()

    ws_m1 = wb_bp.active
    ws_m1.title = "MGR1"
    for r_idx, row in enumerate(dataframe_to_rows(df_m1_print, index=False, header=True), 1):
        for c_idx, val in enumerate(row, 1):
            ws_m1.cell(row=r_idx, column=c_idx, value=val)
    ws_m1.auto_filter.ref = ws_m1.dimensions

    ws_m2 = wb_bp.create_sheet("MGR2")
    for r_idx, row in enumerate(dataframe_to_rows(df_m2_print, index=False, header=True), 1):
        for c_idx, val in enumerate(row, 1):
            ws_m2.cell(row=r_idx, column=c_idx, value=val)
    ws_m2.auto_filter.ref = ws_m2.dimensions

    ws_delimit = wb_bp.create_sheet("DELIMIT")
    ws_delimit.cell(row=1, column=1, value="PANDUAN COPY-PASTE FILTER DYNAMICS AX (PER 14 ITEM)")
    ws_delimit.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True, size=12)

    cur_row = 3
    # M1 Delimit
    so_m1_list = [s for s in dict.fromkeys(df_m1[so_col])]
    ws_delimit.cell(row=cur_row, column=1, value=f"[ MANAGER 1 - TOTAL {len(so_m1_list)} SALES ORDER ]").font = openpyxl.styles.Font(bold=True, color="000080")
    cur_row += 1
    batch_num = 1
    for i in range(0, len(so_m1_list), 14):
        chunk = so_m1_list[i:i+14]
        ws_delimit.cell(row=cur_row, column=1, value=f"Batch M1-{batch_num} ({len(chunk)} item):")
        ws_delimit.cell(row=cur_row, column=2, value=",".join(chunk))
        cur_row += 1
        batch_num += 1

    cur_row += 1
    # M2 Delimit
    so_m2_list = [s for s in dict.fromkeys(df_m2[so_col])]
    ws_delimit.cell(row=cur_row, column=1, value=f"[ MANAGER 2 - TOTAL {len(so_m2_list)} SALES ORDER ]").font = openpyxl.styles.Font(bold=True, color="000080")
    cur_row += 1
    batch_num = 1
    for i in range(0, len(so_m2_list), 14):
        chunk = so_m2_list[i:i+14]
        ws_delimit.cell(row=cur_row, column=1, value=f"Batch M2-{batch_num} ({len(chunk)} item):")
        ws_delimit.cell(row=cur_row, column=2, value=",".join(chunk))
        cur_row += 1
        batch_num += 1

    for ws in [ws_m1, ws_m2, ws_delimit]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col[:40])
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 45)

    safe_save_workbook(wb_bp, bukmas_print_path)
    report_fn(f"File Print selesai dibuat: {bukmas_filename} (M1: {len(df_m1)} baris, M2: {len(df_m2)} baris)")
    return bukmas_print_path

def stage_8_process_do_summary(do_summary_path: Optional[str], report_fn: Callable[[str], None]) -> Optional[str]:
    if not do_summary_path or not os.path.exists(do_summary_path):
        report_fn("Tidak ada file DO Summary External. Tahap ini dilewati.")
        return None
    report_fn(f"Memproses DO Summary: {os.path.basename(do_summary_path)}...")
    return process_do_summary_file(do_summary_path, log_func=report_fn)

def stage_9_sync_examples(files_to_sync: List[Optional[str]], report_fn: Callable[[str], None]):
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
    def __init__(self, work_dir: Optional[str] = None, progress_callback: Optional[Callable[[str], None]] = None, stage_callback: Optional[Callable[[StageMetric], None]] = None):
        self.work_dir_input = work_dir
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
        self.idx_sync = self.timeline.register_stage("STAGE_09_SYNC", "Sinkronisasi Arsip & Output")

    def report(self, message: str, level: str = "INFO"):
        prefix = f"[{level}] " if level != "INFO" else ""
        formatted = f"{prefix}{message}"
        log(formatted)
        if self.progress_callback:
            self.progress_callback(formatted)

    def _execute_stage(self, stage_idx: int, func: Callable, *args, **kwargs) -> Any:
        stage = self.timeline.get_stage(stage_idx)
        self.timeline.start_stage(stage_idx)
        self.report(f"\n==================================================", "STAGE")
        self.report(f"TAHAP [{stage.index}/9]: {stage.name.upper()}", "STAGE")
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
            self.report(f"[OK] Tahap [{stage.index}/9] Selesai ({stage.duration_sec:.2f}s)", "SUCCESS")
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
        self.report("================================================================================")
        self.report("                     MEMULAI PIPELINE HARIAN: DAILY DATA                        ")
        self.report("================================================================================")

        # 1. Resolusi Folder Kerja
        if self.work_dir_input:
            work_dir = resolve_work_dir(self.work_dir_input)
        else:
            work_dir = os.getcwd()
            if not glob.glob(os.path.join(work_dir, "FORMAT POSISISTOK*.xlsx")):
                dated_dirs = list_date_folders(PROJECT_ROOT)
                if dated_dirs:
                    work_dir = dated_dirs[0]

        self.report(f"Direktori Target: {work_dir}")

        # Tahap 1: Deteksi File Input
        inputs: InputBundle = self._execute_stage(self.idx_init, stage_1_discover_inputs, work_dir, self.report)

        # Tahap 2: Pivot STOK
        stok_res: StokBundle = self._execute_stage(self.idx_stok, stage_2_process_stok, inputs.stok_path, self.report)

        # Tahap 3: Pivot TRN
        trn_res: TrnBundle = self._execute_stage(self.idx_trn, stage_3_process_trn, inputs.trn_path, self.report)

        # Tahap 4: Pivot SOOUT
        soout_res: SooutBundle = self._execute_stage(self.idx_so, stage_4_process_soout, inputs.soout_path, self.report)

        # Tahap 5: Pivot POOUT
        poout_res: PooutBundle = self._execute_stage(self.idx_po, stage_4_process_poout if False else stage_5_process_poout, inputs.poout_path, self.report)

        # Tahap 6: Format Cabang
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

        # Tahap 7: Buku Masuk & FIFO (Kondisional)
        if inputs.buku_masuk_path and os.path.exists(inputs.buku_masuk_path):
            bukmas_out = self._execute_stage(self.idx_bm, stage_7_process_buku_masuk_fifo, inputs, self.report)
        else:
            self.timeline.skip_stage(self.idx_bm, "Tidak ada file Buku Masuk hari ini")
            bukmas_out = None
            self.report("\nTahap [7/9] Alokasi Buku Masuk: Dilewati (Tidak ada penerimaan buku hari ini).", "INFO")

        # Tahap 8: DO Summary External (Kondisional)
        if inputs.do_summary_path and os.path.exists(inputs.do_summary_path):
            do_out = self._execute_stage(self.idx_do, stage_8_process_do_summary, inputs.do_summary_path, self.report)
        else:
            self.timeline.skip_stage(self.idx_do, "Tidak ada file DO Summary sore ini")
            do_out = None
            self.report("\nTahap [8/9] DO Summary: Dilewati (File belum diunduh).", "INFO")

        # Tahap 9: Sinkronisasi Contoh
        sync_files = [inputs.format_cabang_path, inputs.soout_path, inputs.poout_path, inputs.trn_path]
        if bukmas_out:
            sync_files.append(bukmas_out)
        if do_out:
            sync_files.append(do_out)

        self._execute_stage(self.idx_sync, stage_9_sync_examples, sync_files, self.report)

        # Timeline Summary
        summary_text = self.timeline.get_summary_text()
        self.report("\n" + summary_text)
        return self.timeline

def run_automation(
    work_dir: Optional[str] = None,
    progress_callback: Optional[Callable[[str], None]] = None,
    stage_callback: Optional[Callable[[StageMetric], None]] = None
) -> ExecutionTimeline:
    pipeline = DailyDataPipeline(
        work_dir=work_dir,
        progress_callback=progress_callback,
        stage_callback=stage_callback
    )
    return pipeline.run()

if __name__ == "__main__":
    target_dir = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        run_automation(work_dir=target_dir)
    except Exception as exc:
        sys.exit(1)
