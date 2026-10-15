import os
import glob
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
import pandas as pd
import openpyxl

APP_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(APP_DIR)

def normalize_item(val: Any) -> str:
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def to_num(series: pd.Series) -> pd.Series:
    clean = series.astype(str).str.replace(',', '').str.strip()
    return pd.to_numeric(clean, errors='coerce').fillna(0)

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

def build_sales_hierarchy(su_file_path: str) -> Dict[str, Dict[str, str]]:
    if not os.path.exists(su_file_path):
        return {}
    df_su = pd.read_excel(su_file_path, sheet_name=0, dtype=str)
    cols_lower = [c.lower() for c in df_su.columns]
    if any("kodsal" in c for c in cols_lower):
        col_kod = [c for c in df_su.columns if "kodsal" in c.lower()][0]
        col_name = [c for c in df_su.columns if "namsal" in c.lower() or "nama" in c.lower()]
        col_gl = [c for c in df_su.columns if "gl" in c.lower()][0]
        col_asm = [c for c in df_su.columns if "asm" in c.lower()][0]
        col_mgr = [c for c in df_su.columns if "mgr" in c.lower()][0]
        su_db = {}
        for _, row in df_su.iterrows():
            kod = str(row[col_kod]).strip()
            su_db[kod] = {
                "GL": str(row[col_gl]).strip() if pd.notna(row[col_gl]) else "",
                "ASM": str(row[col_asm]).strip() if pd.notna(row[col_asm]) else "",
                "MGR": str(row[col_mgr]).strip() if pd.notna(row[col_mgr]) else "",
                "NAME": str(row[col_name[0]]).strip() if col_name and pd.notna(row[col_name[0]]) else ""
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
        su_db[emp] = {"GL": gl, "ASM": asm, "MGR": mgr, "NAME": ""}
    return su_db

@dataclass
class SOOrderItem:
    queue_no: int
    sales_order: str
    created_date: str
    sales_resp: str
    sales_name: str
    gl: str
    asm: str
    mgr: str
    qty_order: float
    allocated_qty: float
    status: str  # "TERPENUHI", "SEBAGIAN", "MENUNGGU STOK"

@dataclass
class ItemInspectionResult:
    item_code: str
    item_title: str
    in_buku_masuk: bool
    buku_masuk_qty: float
    buku_masuk_files: List[str]
    stok_mks: float
    stok_trn: float
    stok_pst: float
    stok_nas: float
    total_so_needed: float
    total_po_needed: float
    total_allocatable: float
    so_orders: List[SOOrderItem]
    fulfillment_summary: str
    formatted_report: str

def inspect_item(query: str, work_dir: str) -> Optional[ItemInspectionResult]:
    """
    Pemeriksaan komprehensif 360 derajat untuk suatu item buku:
    - Status di buku masuk hari ini (beserta kuantiti kiriman)
    - Posisi multi-stok (Makassar, Pusat, Nasional, Transit)
    - Antrean Sales Order (SO) terurut FIFO
    - Simulasi pemenuhan stok (apakah stok cukup untuk memenuhi seluruh SO)
    """
    if not query or not query.strip():
        return None

    clean_query = query.strip()
    ref_dir = os.path.join(PROJECT_ROOT, "REFERENSI")
    su_path = find_file("SALESUNIT2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\", APP_DIR]) or \
              find_file("SU2026.xlsx", [ref_dir, work_dir, PROJECT_ROOT, "D:\\", APP_DIR])
    su_db = build_sales_hierarchy(su_path) if su_path else {}

    # 1. Temukan data item di Format Cabang (untuk kode, judul, dan stok)
    fc_path = find_file("FORMAT POSISISTOK*.xlsx", [work_dir])
    matched_code = None
    matched_title = ""
    stok_pst = 0.0
    stok_nas = 0.0
    stok_mks = 0.0
    stok_trn = 0.0
    total_so_fc = 0.0
    total_po_fc = 0.0

    if fc_path and os.path.exists(fc_path):
        wb_fc = openpyxl.load_workbook(fc_path, data_only=True, read_only=True)
        ws_fc = wb_fc.active
        # Pencarian baris
        for r_idx, row in enumerate(ws_fc.iter_rows(values_only=True), 1):
            if r_idx < 4:
                continue
            if not row or row[0] is None:
                continue
            code_str = normalize_item(row[0])
            title_str = str(row[1] or "").strip()

            # Match kode atau substring judul
            is_match = False
            if clean_query.lower() in code_str.lower():
                is_match = True
            elif clean_query.lower() in title_str.lower():
                is_match = True
            elif clean_query.lstrip('0') == code_str.lstrip('0') and clean_query.lstrip('0') != "":
                is_match = True

            if is_match:
                matched_code = code_str
                matched_title = title_str
                stok_pst = float(row[32] or 0)
                stok_nas = float(row[33] or 0)
                stok_mks = float(row[34] or 0)
                stok_trn = float(row[35] or 0)
                total_so_fc = float(row[36] or 0)
                total_po_fc = float(row[37] or 0)
                break

    # Jika belum ketemu di Format Cabang, coba cari di SOOUT
    soout_path = find_file("SOOUT*.xlsx", [work_dir])
    if not matched_code and soout_path and os.path.exists(soout_path):
        df_so_peek = pd.read_excel(soout_path, sheet_name="Sheet1", dtype=str)
        it_cols = [c for c in df_so_peek.columns if "item number" in c.lower()]
        name_cols = [c for c in df_so_peek.columns if "item name" in c.lower()]
        if it_cols:
            it_col = it_cols[0]
            name_col = name_cols[0] if name_cols else it_col
            for _, r in df_so_peek.iterrows():
                c_str = normalize_item(r[it_col])
                t_str = str(r[name_col] or "").strip()
                if (clean_query.lower() in c_str.lower()) or (clean_query.lower() in t_str.lower()):
                    matched_code = c_str
                    matched_title = t_str
                    break

    # 2. Cek di File Buku Masuk Hari Ini
    bm_files = []
    for pat in ["buku masuk*.xlsx", "BUKMAS*.xlsx"]:
        for f in glob.glob(os.path.join(work_dir, pat)):
            if "PRINT" not in os.path.basename(f).upper() and f not in bm_files:
                bm_files.append(f)

    in_buku_masuk = False
    buku_masuk_qty = 0.0
    matched_bm_files = []

    # Jika belum ketemu di FC dan SOOUT, cek apakah ada di Buku Masuk
    if not matched_code:
        for bm_f in bm_files:
            try:
                df_bm = pd.read_excel(bm_f, dtype=str)
                it_cols = [c for c in df_bm.columns if "item number" in c.lower()]
                t_cols = [c for c in df_bm.columns if "title" in c.lower() or "judul" in c.lower() or "name" in c.lower()]
                if it_cols:
                    it_col = it_cols[0]
                    t_col = t_cols[0] if t_cols else None
                    for _, r in df_bm.iterrows():
                        c_str = normalize_item(r[it_col])
                        t_str = str(r[t_col] or "").strip() if t_col else ""
                        if (clean_query.lower() in c_str.lower()) or (clean_query.lower() in t_str.lower()):
                            matched_code = c_str
                            matched_title = t_str or "[Item Baru di Buku Masuk]"
                            break
                    if matched_code:
                        break
            except Exception:
                pass

    # Jika tetap tidak ditemukan di mana pun (FC, SOOUT, BUKMAS), maka item benar-benar tidak ada
    if not matched_code:
        return None

    # Hitung kuantiti di Buku Masuk
    for bm_f in bm_files:
        try:
            df_bm = pd.read_excel(bm_f, dtype=str)
            it_cols = [c for c in df_bm.columns if "item number" in c.lower()]
            qty_cols = [c for c in df_bm.columns if "quantity" in c.lower() or "qty" in c.lower()]
            if it_cols:
                it_col = it_cols[0]
                df_bm['it_clean'] = df_bm[it_col].apply(normalize_item)
                sub_bm = df_bm[df_bm['it_clean'] == matched_code]
                if not sub_bm.empty:
                    in_buku_masuk = True
                    matched_bm_files.append(os.path.basename(bm_f))
                    if qty_cols:
                        qty_col = qty_cols[0]
                        buku_masuk_qty += to_num(sub_bm[qty_col]).sum()
                    else:
                        buku_masuk_qty += len(sub_bm)
        except Exception:
            pass

    # 3. Ambil Pesanan SO (SOOUT) & Susun FIFO Queue
    so_orders: List[SOOrderItem] = []
    total_so_needed = 0.0

    if soout_path and os.path.exists(soout_path):
        df_so_raw = pd.read_excel(soout_path, sheet_name="Sheet1", dtype=str)
        so_col = [c for c in df_so_raw.columns if "sales order" in c.lower()][0]
        date_col = [c for c in df_so_raw.columns if "created date" in c.lower()][0]
        it_col = [c for c in df_so_raw.columns if "item number" in c.lower()][0]
        resp_col = [c for c in df_so_raw.columns if "sales responsible" in c.lower()][0]
        rem_col = [c for c in df_so_raw.columns if "deliver reminder" in c.lower() or "deliver remainder" in c.lower()][0]
        name_cols = [c for c in df_so_raw.columns if "sales resp. name" in c.lower() or "name" in c.lower()]
        resp_name_col = name_cols[0] if name_cols else None

        df_so_raw['it_clean'] = df_so_raw[it_col].apply(normalize_item)
        sub_so = df_so_raw[(df_so_raw['it_clean'] == matched_code) &
                           (df_so_raw[so_col].astype(str).str.startswith("56-SOE"))].copy()

        if not sub_so.empty:
            sub_so['qty_order'] = to_num(sub_so[rem_col]).abs()
            sub_so = sub_so[sub_so['qty_order'] > 0]
            sub_so['dt_sort'] = pd.to_datetime(sub_so[date_col], format='mixed', dayfirst=True, errors='coerce')
            sub_so = sub_so.sort_values(by=['dt_sort', so_col], ascending=[True, True])

            total_so_needed = sub_so['qty_order'].sum()

            # Simulasi Alokasi Stok FIFO
            total_allocatable = stok_mks + buku_masuk_qty
            current_available = total_allocatable

            q_no = 1
            for _, r in sub_so.iterrows():
                q_val = float(r['qty_order'])
                resp = str(r[resp_col]).strip()
                h_info = su_db.get(resp, {"GL": "", "ASM": "", "MGR": "", "NAME": ""})
                s_name = str(r[resp_name_col]).strip() if resp_name_col and pd.notna(r[resp_name_col]) else h_info.get("NAME", "")

                allocated = 0.0
                if current_available >= q_val:
                    allocated = q_val
                    current_available -= q_val
                    st_str = "TERPENUHI"
                elif current_available > 0:
                    allocated = current_available
                    current_available = 0.0
                    st_str = f"SEBAGIAN ({allocated:,.0f}/{q_val:,.0f})"
                else:
                    allocated = 0.0
                    st_str = "MENUNGGU STOK"

                date_display = str(r[date_col])[:16] if pd.notna(r[date_col]) else "-"
                item_obj = SOOrderItem(
                    queue_no=q_no,
                    sales_order=str(r[so_col]).strip(),
                    created_date=date_display,
                    sales_resp=resp,
                    sales_name=s_name,
                    gl=h_info.get("GL", "-"),
                    asm=h_info.get("ASM", "-"),
                    mgr=h_info.get("MGR", "-"),
                    qty_order=q_val,
                    allocated_qty=allocated,
                    status=st_str
                )
                so_orders.append(item_obj)
                q_no += 1
    else:
        total_so_needed = total_so_fc

    total_allocatable = stok_mks + buku_masuk_qty

    # Summary Text Pemenuhan
    if not so_orders:
        if total_so_needed == 0:
            fulfillment_summary = "Tidak ada pesanan SO terbuka (Deliver Reminder = 0)."
        else:
            fulfillment_summary = f"Total kebutuhan SO: {total_so_needed:,.0f} eks."
    else:
        fulfilled_count = sum(1 for o in so_orders if o.status == "TERPENUHI")
        if fulfilled_count == len(so_orders):
            surplus = total_allocatable - total_so_needed
            fulfillment_summary = f"Semua {len(so_orders)} pesanan SO TERPENUHI 100%! (Surplus stok siap: {surplus:,.0f} eks)"
        elif any("SEBAGIAN" in o.status for o in so_orders) or fulfilled_count > 0:
            shortage = total_so_needed - total_allocatable
            fulfillment_summary = f"Terpenuhi sebagian ({fulfilled_count}/{len(so_orders)} SO). Masih kurang: {shortage:,.0f} eks."
        else:
            fulfillment_summary = f"Belum dapat terpenuhi (Stok Makassar + Buku Masuk = {total_allocatable:,.0f} eks, Butuh: {total_so_needed:,.0f} eks)."

    # Format Teks Laporan
    rep_lines = [
        "================================================================================",
        f"        HASIL PEMERIKSAAN ITEM & ANTREAN SO FIFO: {matched_code}       ",
        "================================================================================",
        f" Kode Buku   : {matched_code}",
        f" Judul Buku  : {matched_title}",
        "--------------------------------------------------------------------------------",
        " [STATUS BUKU MASUK HARIAN]",
    ]
    if in_buku_masuk:
        files_str = ", ".join(matched_bm_files)
        rep_lines.append(f"  Status     : TERDAFTAR DI BUKU MASUK HARI INI")
        rep_lines.append(f"  Qty Masuk  : {buku_masuk_qty:,.0f} eksemplar (dari {files_str})")
    else:
        rep_lines.append(f"  Status     : BELUM ADA DI BUKU MASUK HARI INI")

    rep_lines.extend([
        "--------------------------------------------------------------------------------",
        " [POSISI MULTI-STOK]",
        f"  Stok Fisik Makassar (56) : {stok_mks:>8,.0f} eks",
        f"  Sedang Transit (TRN)     : {stok_trn:>8,.0f} eks",
        f"  Stok Gudang Pusat (50)   : {stok_pst:>8,.0f} eks",
        f"  Stok Nasional            : {stok_nas:>8,.0f} eks",
        f"  Total SO Out Terbuka     : {total_so_needed:>8,.0f} eks",
        f"  PO Out                   : {total_po_fc:>8,.0f} eks",
        "--------------------------------------------------------------------------------",
        f" [KAPASITAS PEMENUHAN STOK]",
        f"  Stok Siap Alokasi (MKS + BM) : {total_allocatable:>6,.0f} eks",
        f"  Total Kebutuhan SO           : {total_so_needed:>6,.0f} eks",
        f"  Status                       : {fulfillment_summary}",
        "--------------------------------------------------------------------------------",
        " [ANTREAN SALES ORDER (FIFO / URUT TANGGAL)]",
    ])

    if not so_orders:
        rep_lines.append("  (Tidak ada pesanan SO 56-SOE terbuka untuk item ini)")
    else:
        rep_lines.append(f"  {'No':<3} | {'Sales Order':<18} | {'Tgl Order':<16} | {'Qty':<5} | {'Pemesan':<24} | {'Manager':<16} | {'Status'}")
        rep_lines.append("  " + "-" * 105)
        for o in so_orders:
            sales_info = f"{o.sales_resp} ({o.gl})"
            if len(sales_info) > 24:
                sales_info = sales_info[:21] + "..."
            rep_lines.append(
                f"  #{o.queue_no:<2} | {o.sales_order:<18} | {o.created_date:<16} | {o.qty_order:>5,.0f} | {sales_info:<24} | {o.mgr:<16} | {o.status}"
            )

    rep_lines.append("================================================================================")
    formatted_report = "\n".join(rep_lines)

    return ItemInspectionResult(
        item_code=matched_code,
        item_title=matched_title,
        in_buku_masuk=in_buku_masuk,
        buku_masuk_qty=buku_masuk_qty,
        buku_masuk_files=matched_bm_files,
        stok_mks=stok_mks,
        stok_trn=stok_trn,
        stok_pst=stok_pst,
        stok_nas=stok_nas,
        total_so_needed=total_so_needed,
        total_po_needed=total_po_fc,
        total_allocatable=total_allocatable,
        so_orders=so_orders,
        fulfillment_summary=fulfillment_summary,
        formatted_report=formatted_report
    )
