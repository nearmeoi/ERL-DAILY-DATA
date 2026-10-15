import os
import glob
import re
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Callable
from datetime import datetime, timedelta
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(APP_DIR)

def normalize_item(val: Any) -> str:
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.zfill(10) if s else ""

def parse_folder_date(folder_name: str) -> Optional[datetime]:
    base = os.path.basename(folder_name)
    m = re.search(r'(\d{1,2})\s*([A-Za-z]+)\s*(\d{4})', base)
    if not m:
        return None
    day = int(m.group(1))
    mon_str = m.group(2).upper()
    year = int(m.group(3))
    months = {
        "JAN": 1, "JANUARI": 1, "FEB": 2, "FEBRUARI": 2, "MAR": 3, "MARET": 3,
        "APR": 4, "APRIL": 4, "MEI": 5, "MAY": 5, "JUN": 6, "JUNI": 6,
        "JUL": 7, "JULI": 7, "AGU": 8, "AGUSTUS": 8, "AUG": 8, "SEP": 9, "SEPTEMBER": 9,
        "OKT": 10, "OKTOBER": 10, "OCT": 10, "OCTOBER": 10, "NOV": 11, "NOVEMBER": 11,
        "DES": 12, "DESEMBER": 12, "DEC": 12, "DECEMBER": 12
    }
    month = months.get(mon_str)
    if not month:
        return None
    try:
        return datetime(year, month, day)
    except ValueError:
        return None

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

def build_sales_hierarchy(su_path: str) -> Dict[str, Dict[str, str]]:
    if not su_path or not os.path.exists(su_path):
        return {}
    try:
        df_su = pd.read_excel(su_path, sheet_name=0, dtype=str)
    except Exception:
        return {}
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
            u_upper = desc.upper()
            if "GL " in u_upper or u_upper.startswith("GL"):
                gl = desc
            elif "ASM " in u_upper or u_upper.startswith("ASM"):
                asm = desc
            elif "MGR " in u_upper or "MANAGER" in u_upper:
                mgr = desc
            curr = unit_to_parent.get(curr, "")
        su_db[emp] = {"GL": gl, "ASM": asm, "MGR": mgr}
    return su_db

@dataclass
class RecheckItem:
    sales_order: str
    surat_pesanan: str
    created_date: str
    customer_account: str
    item_number: str
    item_name: str
    qty_ordered: float
    sisa_kirim: float
    stok_fisik_wh56: float
    onhand_sistem: float
    mgr: str
    asm: str
    gl: str
    status_diagnosa: str
    keterangan: str
    sumber_alokasi: str  # 'SP Kemarin di Meja' atau 'Alokasi Buku Masuk Kemarin'

@dataclass
class RecheckSPResult:
    work_dir: str
    target_date_str: str
    prev_date_str: str
    total_checked: int
    ready_to_post: List[RecheckItem]
    potential_diff: List[RecheckItem]
    waiting_stock: List[RecheckItem]
    all_items: List[RecheckItem]
    delimit_batches_mgr1: List[str]
    delimit_batches_mgr2: List[str]
    delimit_batches_all: List[str]
    excel_output_path: Optional[str] = None


def run_recheck_sp(
    work_dir: str,
    progress_fn: Optional[Callable[[str], None]] = None
) -> RecheckSPResult:
    def log(msg: str):
        if progress_fn:
            progress_fn(msg)
        else:
            try:
                print(msg, flush=True)
            except UnicodeEncodeError:
                safe = msg.encode("ascii", errors="replace").decode("ascii")
                print(safe, flush=True)

    log("=" * 60)
    log("MEMULAI PROSES RECHECK SP KEMARIN & VALIDASI POSTING")
    log("=" * 60)

    # 1. Deteksi Tanggal Kerja & Tanggal Kemarin
    curr_dt = parse_folder_date(work_dir)
    if not curr_dt:
        curr_dt = datetime.now()
    date_str = curr_dt.strftime("%d %b %Y").upper()

    # Cari daftar folder tanggal untuk menemukan folder tanggal sebelumnya
    all_date_dirs = []
    parent_dir = os.path.dirname(os.path.abspath(work_dir))
    for d in glob.glob(os.path.join(parent_dir, "*")):
        if os.path.isdir(d) and parse_folder_date(d):
            all_date_dirs.append((parse_folder_date(d), d))
    all_date_dirs.sort(key=lambda x: x[0])

    prev_folder = None
    for idx, (p_dt, p_path) in enumerate(all_date_dirs):
        if os.path.abspath(p_path) == os.path.abspath(work_dir):
            if idx > 0:
                prev_folder = all_date_dirs[idx - 1][1]
            break

    # Tentukan tanggal acuan kemarin (H-1 atau Jumat-Minggu jika Senin)
    if curr_dt.weekday() == 0:  # Hari Senin
        target_dates_str = {
            (curr_dt - timedelta(days=1)).strftime("%Y-%m-%d"),
            (curr_dt - timedelta(days=2)).strftime("%Y-%m-%d"),
            (curr_dt - timedelta(days=3)).strftime("%Y-%m-%d"),
        }
        prev_date_label = f"{(curr_dt - timedelta(days=3)).strftime('%d')} - {(curr_dt - timedelta(days=1)).strftime('%d %b %Y').upper()}"
    else:
        prev_dt = curr_dt - timedelta(days=1)
        target_dates_str = {prev_dt.strftime("%Y-%m-%d")}
        prev_date_label = prev_dt.strftime("%d %b %Y").upper()

    log(f"Tanggal Operasional Hari Ini : {date_str}")
    log(f"Tanggal Acuan Recheck Kemarin: {prev_date_label}")
    if prev_folder:
        log(f"Folder Referensi Kemarin     : {os.path.basename(prev_folder)}")

    # 2. File Resolution
    soout_candidates = glob.glob(os.path.join(work_dir, "SOOUT*.xlsx"))
    stok_candidates = glob.glob(os.path.join(work_dir, "STOK*.xlsx"))
    if not soout_candidates:
        raise FileNotFoundError(f"File SOOUT tidak ditemukan di folder '{work_dir}'.")
    if not stok_candidates:
        raise FileNotFoundError(f"File STOK tidak ditemukan di folder '{work_dir}'.")

    soout_path = soout_candidates[0]
    stok_path = stok_candidates[0]
    log(f"Membaca File SOOUT Pagi Ini  : {os.path.basename(soout_path)}")
    log(f"Membaca File STOK Pagi Ini   : {os.path.basename(stok_path)}")

    # Master Sales SU2026
    ref_dir = os.path.join(PROJECT_ROOT, "REFERENSI")
    su_path = find_file("SALESUNIT2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\"]) or \
              find_file("SU2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\"])
    su_db = build_sales_hierarchy(su_path) if su_path else {}
    if su_db:
        log(f"Master Sales Dimuat          : {len(su_db):,} sales responsible")

    # 3. Baca STOK Fisik Makassar (WH 56)
    log("Memproses stok fisik Makassar (Warehouse 56)...")
    df_stk = pd.read_excel(stok_path, sheet_name="Sheet1")
    wh_col = [c for c in df_stk.columns if "warehouse" in c.lower()][0]
    it_col = [c for c in df_stk.columns if "item number" in c.lower()][0]
    phys_col = [c for c in df_stk.columns if "available physical" in c.lower()][0]

    stk_56 = df_stk[df_stk[wh_col].astype(str).str.strip() == "56"].copy()
    stk_56['clean_it'] = stk_56[it_col].apply(normalize_item)
    wh56_stock = stk_56.groupby('clean_it')[phys_col].sum().to_dict()
    log(f"Total item dengan stok fisik di WH 56: {len(wh56_stock):,} item.")

    # 4. Baca SOOUT Hari Ini
    log("Menganalisis pesanan berstatus terbuka di SOOUT...")
    df_so = pd.read_excel(soout_path, sheet_name="Sheet1")
    so_col = [c for c in df_so.columns if "sales order" in c.lower()][0]
    sp_col = [c for c in df_so.columns if "surat pesanan" in c.lower()][0]
    dt_col = [c for c in df_so.columns if "created date" in c.lower()][0]
    it_so_col = [c for c in df_so.columns if "item number" in c.lower()][0]
    nm_col = [c for c in df_so.columns if "item name" in c.lower()][0]
    rem_col = [c for c in df_so.columns if "deliver reminder" in c.lower()][0]
    ord_col = [c for c in df_so.columns if "qty ordered" in c.lower()][0]
    cust_col = [c for c in df_so.columns if "customer account" in c.lower()][0]
    onh_col = [c for c in df_so.columns if "on-hand" in c.lower()][0]
    resp_col = [c for c in df_so.columns if "sales responsible" in c.lower()][0]

    df_soe = df_so[df_so[so_col].astype(str).str.startswith("56-SOE")].copy()
    df_soe['clean_it'] = df_soe[it_so_col].apply(normalize_item)
    df_soe['sisa'] = pd.to_numeric(df_soe[rem_col], errors='coerce').fillna(0).abs()
    df_soe['dt_obj'] = pd.to_datetime(df_soe[dt_col], format='mixed', dayfirst=True, errors='coerce')
    df_soe['dt_str'] = df_soe['dt_obj'].dt.strftime("%Y-%m-%d")
    df_soe['onh_num'] = pd.to_numeric(df_soe[onh_col], errors='coerce').fillna(0)
    df_soe['stok_fisik'] = df_soe['clean_it'].map(wh56_stock).fillna(0)

    # 5. Kumpulkan Target SO Kemarin
    # Kelompok A: SO yang tanggal pembuatannya kemarin
    so_kemarin_df = df_soe[df_soe['dt_str'].isin(target_dates_str) & (df_soe['sisa'] > 0)].copy()

    # Kelompok B: SO dari BUKMAS PRINT kemarin yang belum terbit DO
    prev_bm_print_sos = set()
    prev_do_sos = set()
    if prev_folder and os.path.exists(prev_folder):
        bm_print_files = glob.glob(os.path.join(prev_folder, "BUKMAS*PRINT*.xlsx"))
        if bm_print_files:
            try:
                for sname in ["MGR1", "MGR2"]:
                    df_p = pd.read_excel(bm_print_files[0], sheet_name=sname)
                    if "Sales order" in df_p.columns:
                        prev_bm_print_sos.update(df_p["Sales order"].dropna().astype(str).str.strip().unique())
            except Exception:
                pass
        do_files = glob.glob(os.path.join(prev_folder, "Do_Summary_External*.xlsx")) + \
                   glob.glob(os.path.join(prev_folder, "Do_Summary_External*.csv"))
        if do_files:
            try:
                df_do = pd.read_excel(do_files[0]) if do_files[0].endswith('.xlsx') else pd.read_csv(do_files[0])
                sales_cols = [c for c in df_do.columns if "sales" in c.lower() or "so" in c.lower()]
                if sales_cols:
                    prev_do_sos.update(df_do[sales_cols[0]].dropna().astype(str).str.strip().unique())
            except Exception:
                pass

    bm_missed_sos = prev_bm_print_sos - prev_do_sos
    so_bm_missed_df = df_soe[df_soe[so_col].astype(str).isin(bm_missed_sos) & (df_soe['sisa'] > 0)].copy()

    # Gabungkan kandidat
    combined_records = []
    seen_keys = set()

    # Masukkan Kelompok A
    for _, r in so_kemarin_df.iterrows():
        key = (str(r[so_col]).strip(), str(r['clean_it']).strip())
        if key not in seen_keys:
            seen_keys.add(key)
            combined_records.append((r, "SP Kemarin di Meja"))

    # Masukkan Kelompok B
    for _, r in so_bm_missed_df.iterrows():
        key = (str(r[so_col]).strip(), str(r['clean_it']).strip())
        if key not in seen_keys:
            seen_keys.add(key)
            combined_records.append((r, "Alokasi Buku Masuk Kemarin"))

    log(f"Total baris transaksi kemarin dievaluasi: {len(combined_records):,} baris.")

    # 6. Analisis dan Klasifikasi Diagnosa
    ready_to_post: List[RecheckItem] = []
    potential_diff: List[RecheckItem] = []
    waiting_stock: List[RecheckItem] = []
    all_items: List[RecheckItem] = []

    # Kumpulkan item yang baru masuk fisik di Buku Masuk (kemarin atau hari ini)
    fresh_bm_items = set()
    for fld in ([prev_folder] if prev_folder else []) + [work_dir]:
        if not fld or not os.path.exists(fld):
            continue
        for bmp in glob.glob(os.path.join(fld, "buku masuk*.xlsx")) + glob.glob(os.path.join(fld, "BUKMAS*.xlsx")):
            if "PRINT" not in os.path.basename(bmp).upper():
                try:
                    df_b_raw = pd.read_excel(bmp, dtype=str)
                    b_cols = [c for c in df_b_raw.columns if "item number" in c.lower()]
                    if b_cols:
                        fresh_bm_items.update({normalize_item(x) for x in df_b_raw[b_cols[0]].dropna().unique() if normalize_item(x)})
                except Exception:
                    pass

    # Deteksi item yang memiliki pesanan lama menggantung (> 14 hari) padahal ada stok di sistem
    cutoff_dt = curr_dt - timedelta(days=14)
    old_stuck_items = set()
    old_sub = df_soe[(df_soe['dt_obj'] < cutoff_dt) & (df_soe['sisa'] > 0)]
    for _, o_row in old_sub.iterrows():
        it_c = str(o_row['clean_it'])
        if it_c not in fresh_bm_items:
            old_stuck_items.add(it_c)

    for row, sumber in combined_records:
        so = str(row[so_col]).strip()
        sp = str(row[sp_col]).strip() if pd.notna(row[sp_col]) else "-"
        tgl_display = str(row[dt_col])[:16] if pd.notna(row[dt_col]) else "-"
        cust = str(row[cust_col]).strip() if pd.notna(row[cust_col]) else "-"
        it_num = str(row['clean_it'])
        it_name = str(row[nm_col]).strip() if pd.notna(row[nm_col]) else "-"
        q_order = float(row[ord_col]) if pd.notna(row[ord_col]) else 0.0
        sisa = float(row['sisa'])
        stok_fisik = float(row['stok_fisik'])
        onh = float(row['onh_num'])
        resp = str(row[resp_col]).strip() if pd.notna(row[resp_col]) else ""

        h_info = su_db.get(resp, {"GL": "-", "ASM": "-", "MGR": "-"})
        mgr = h_info.get("MGR", "-")
        asm = h_info.get("ASM", "-")
        gl = h_info.get("GL", "-")

        is_fresh_receipt = (it_num in fresh_bm_items)
        is_chronic_ghost = (it_num in old_stuck_items and not is_fresh_receipt)

        if stok_fisik >= sisa and stok_fisik > 0 and not is_chronic_ghost:
            status = "SIAP POSTING (TERLEWAT)"
            ket = f"Stok fisik Makassar cukup ({stok_fisik:,.0f} eks), barang siap diposting"
        elif is_chronic_ghost or (onh >= sisa and stok_fisik <= 0):
            status = "POTENSI SELISIH FISIK"
            if is_chronic_ghost:
                ket = f"Stok sistem ada ({stok_fisik if stok_fisik>0 else onh:,.0f} eks), tapi pesanan lama >14 hari mandek. Kemungkinan fisik kosong di rak"
            else:
                ket = f"On-hand sistem ada ({onh:,.0f} eks), tapi fisik WH 56 kosong (0 eks) - perlu opname"
        else:
            status = "MENUNGGU STOK"
            ket = f"Stok fisik WH 56 belum mencukupi (sisa: {sisa:,.0f} eks, fisik: {stok_fisik:,.0f} eks)"

        item_obj = RecheckItem(
            sales_order=so,
            surat_pesanan=sp,
            created_date=tgl_display,
            customer_account=cust,
            item_number=it_num,
            item_name=it_name,
            qty_ordered=q_order,
            sisa_kirim=sisa,
            stok_fisik_wh56=stok_fisik,
            onhand_sistem=onh,
            mgr=mgr,
            asm=asm,
            gl=gl,
            status_diagnosa=status,
            keterangan=ket,
            sumber_alokasi=sumber
        )
        all_items.append(item_obj)
        if status == "SIAP POSTING (TERLEWAT)":
            ready_to_post.append(item_obj)
        elif status == "POTENSI SELISIH FISIK":
            potential_diff.append(item_obj)
        else:
            waiting_stock.append(item_obj)

    # 7. Susun Batch Delimit 14 Item untuk SO yang SIAP POSTING
    ready_sos_m1 = []
    ready_sos_m2 = []
    ready_sos_all = []
    seen_so_ready = set()

    for item in ready_to_post:
        if item.sales_order not in seen_so_ready:
            seen_so_ready.add(item.sales_order)
            ready_sos_all.append(item.sales_order)
            if "1" in str(item.mgr):
                ready_sos_m1.append(item.sales_order)
            elif "2" in str(item.mgr):
                ready_sos_m2.append(item.sales_order)
            else:
                ready_sos_m1.append(item.sales_order)

    def make_batches(so_list: List[str]) -> List[str]:
        batches = []
        for i in range(0, len(so_list), 14):
            batches.append(",".join(so_list[i:i+14]))
        return batches

    delimit_m1 = make_batches(ready_sos_m1)
    delimit_m2 = make_batches(ready_sos_m2)
    delimit_all = make_batches(ready_sos_all)

    log("-" * 60)
    log(f"HASIL RECHECK SP KEMARIN:")
    log(f"  - Total Transaksi Kemarin Diperiksa : {len(all_items):,} baris ({len({x.sales_order for x in all_items}):,} SO)")
    log(f"  - [SIAP POSTING] (TERLEWAT)         : {len(ready_to_post):,} baris ({len(ready_sos_all):,} SO)")
    log(f"  - [POTENSI SELISIH STOK FISIK]      : {len(potential_diff):,} baris ({len({x.sales_order for x in potential_diff}):,} SO)")
    log(f"  - [MENUNGGU STOK] (INDENT)          : {len(waiting_stock):,} baris ({len({x.sales_order for x in waiting_stock}):,} SO)")
    log("-" * 60)

    # 8. Tulis File Excel Laporan: RECHECK_SP_<TANGGAL>.xlsx
    out_filename = f"RECHECK_SP_{date_str}.xlsx"
    out_path = os.path.join(work_dir, out_filename)

    wb = openpyxl.Workbook()
    # Sheet 1: SIAP POSTING (TERLEWAT)
    ws1 = wb.active
    ws1.title = "SIAP_POSTING"

    # Styling helper
    font_head = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_blue = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_yellow = PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")
    fill_gray = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers1 = [
        "No", "Sales Order", "Surat Pesanan", "MGR", "ASM", "GL",
        "Tanggal Dibuat", "Customer Account", "Item Number", "Judul Buku",
        "Qty Pesan", "Sisa Kirim", "Stok Fisik WH 56", "Sumber Alokasi"
    ]

    for col_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col_idx, value=h)
        cell.font = font_head
        cell.fill = fill_blue
        cell.alignment = align_center

    for r_idx, itm in enumerate(ready_to_post, 2):
        row_vals = [
            r_idx - 1, itm.sales_order, itm.surat_pesanan, itm.mgr, itm.asm, itm.gl,
            itm.created_date, itm.customer_account, itm.item_number, itm.item_name,
            itm.qty_ordered, itm.sisa_kirim, itm.stok_fisik_wh56, itm.sumber_alokasi
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws1.cell(row=r_idx, column=c_idx, value=val)
            c.border = thin_border
            if c_idx in [1, 2, 3, 7, 8, 9]:
                c.alignment = align_center
            elif c_idx in [11, 12, 13]:
                c.alignment = align_right
                c.number_format = "#,##0"
            else:
                c.alignment = align_left

    ws1.auto_filter.ref = ws1.dimensions

    # Sheet 2: DELIMIT (Batch 14 Item Siap Posting)
    ws2 = wb.create_sheet("DELIMIT")
    ws2.cell(row=1, column=1, value="PANDUAN COPY-PASTE FILTER DYNAMICS AX (PER 14 ITEM SIAP POSTING)").font = Font(size=12, bold=True, color="1E3A8A")

    cur_row = 3
    if delimit_m1:
        ws2.cell(row=cur_row, column=1, value=f"[ MANAGER 1 - TOTAL {len(ready_sos_m1)} SALES ORDER SIAP POSTING ]").font = Font(bold=True, color="000080")
        cur_row += 1
        for b_i, b_str in enumerate(delimit_m1, 1):
            ws2.cell(row=cur_row, column=1, value=f"Batch M1-{b_i} ({len(b_str.split(','))} SO):")
            ws2.cell(row=cur_row, column=2, value=b_str)
            cur_row += 1
        cur_row += 1

    if delimit_m2:
        ws2.cell(row=cur_row, column=1, value=f"[ MANAGER 2 - TOTAL {len(ready_sos_m2)} SALES ORDER SIAP POSTING ]").font = Font(bold=True, color="000080")
        cur_row += 1
        for b_i, b_str in enumerate(delimit_m2, 1):
            ws2.cell(row=cur_row, column=1, value=f"Batch M2-{b_i} ({len(b_str.split(','))} SO):")
            ws2.cell(row=cur_row, column=2, value=b_str)
            cur_row += 1
        cur_row += 1

    if not ready_to_post:
        ws2.cell(row=cur_row, column=1, value="Tidak ada Sales Order yang siap diposting (Semua terposting / menunggu stok).")

    # Sheet 3: POTENSI SELISIH FISIK
    ws3 = wb.create_sheet("POTENSI_SELISIH_STOK")
    headers3 = [
        "No", "Sales Order", "Surat Pesanan", "MGR", "Customer Account",
        "Item Number", "Judul Buku", "Sisa Kirim", "On-hand Sistem", "Stok Fisik WH 56", "Diagnosa Masalah"
    ]
    for col_idx, h in enumerate(headers3, 1):
        cell = ws3.cell(row=1, column=col_idx, value=h)
        cell.font = font_head
        cell.fill = fill_yellow
        cell.alignment = align_center

    for r_idx, itm in enumerate(potential_diff, 2):
        row_vals = [
            r_idx - 1, itm.sales_order, itm.surat_pesanan, itm.mgr, itm.customer_account,
            itm.item_number, itm.item_name, itm.sisa_kirim, itm.onhand_sistem, itm.stok_fisik_wh56, itm.keterangan
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws3.cell(row=r_idx, column=c_idx, value=val)
            c.border = thin_border
            if c_idx in [1, 2, 3, 5, 6]:
                c.alignment = align_center
            elif c_idx in [8, 9, 10]:
                c.alignment = align_right
                c.number_format = "#,##0"
            else:
                c.alignment = align_left

    ws3.auto_filter.ref = ws3.dimensions

    # Sheet 4: SEMUA TRANSAKSI DIEVALUASI
    ws4 = wb.create_sheet("SEMUA_RECHECK")
    headers4 = [
        "No", "Sales Order", "Surat Pesanan", "MGR", "Tanggal Dibuat", "Customer Account",
        "Item Number", "Judul Buku", "Qty Pesan", "Sisa Kirim", "Stok Fisik WH 56", "On-hand Sistem",
        "Status Diagnosa", "Keterangan", "Sumber Alokasi"
    ]
    for col_idx, h in enumerate(headers4, 1):
        cell = ws4.cell(row=1, column=col_idx, value=h)
        cell.font = font_head
        cell.fill = fill_gray
        cell.alignment = align_center

    for r_idx, itm in enumerate(all_items, 2):
        row_vals = [
            r_idx - 1, itm.sales_order, itm.surat_pesanan, itm.mgr, itm.created_date, itm.customer_account,
            itm.item_number, itm.item_name, itm.qty_ordered, itm.sisa_kirim, itm.stok_fisik_wh56, itm.onhand_sistem,
            itm.status_diagnosa, itm.keterangan, itm.sumber_alokasi
        ]
        for c_idx, val in enumerate(row_vals, 1):
            c = ws4.cell(row=r_idx, column=c_idx, value=val)
            c.border = thin_border
            if c_idx in [1, 2, 3, 5, 6, 7]:
                c.alignment = align_center
            elif c_idx in [9, 10, 11, 12]:
                c.alignment = align_right
                c.number_format = "#,##0"
            else:
                c.alignment = align_left

    ws4.auto_filter.ref = ws4.dimensions

    # Auto-adjust column widths
    for sheet in [ws1, ws2, ws3, ws4]:
        for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if "\n" in val_str:
                    val_str = val_str.split("\n")[0]
                if len(val_str) > max_len:
                    max_len = len(val_str)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

    wb.save(out_path)
    log(f"[SUKSES] File Laporan Recheck tersimpan: {os.path.basename(out_path)}")

    return RecheckSPResult(
        work_dir=work_dir,
        target_date_str=date_str,
        prev_date_str=prev_date_label,
        total_checked=len(all_items),
        ready_to_post=ready_to_post,
        potential_diff=potential_diff,
        waiting_stock=waiting_stock,
        all_items=all_items,
        delimit_batches_mgr1=delimit_m1,
        delimit_batches_mgr2=delimit_m2,
        delimit_batches_all=delimit_all,
        excel_output_path=out_path
    )
