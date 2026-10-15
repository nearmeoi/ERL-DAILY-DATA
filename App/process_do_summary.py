import os
import sys
import glob
import re
import subprocess
import shutil
import datetime
from typing import Optional, List, Tuple, Callable

# Template PowerShell script untuk eksekusi Excel COM native
PS_SCRIPT_TEMPLATE = r'''
param(
    [string]$csvPath,
    [string]$outXlsx,
    [string]$delimiter = ","
)

$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 1. Buka file CSV
    $wb = $excel.Workbooks.Open($csvPath)
    $rawSheet = $wb.ActiveSheet

    # 2. Text to Columns pada Kolom A jika data masih menumpuk di Kolom A
    $valB1 = $rawSheet.Cells(1, 2).Value2
    if (-not $valB1) {
        $colA = $rawSheet.Range("A:A")
        $isSemi = ($delimiter -eq ";")
        # DataType: 1=xlDelimited, ConsecutiveDelimiter:$false, Tab:$false, Semicolon:$isSemi, Comma:(-not $isSemi), Space:$false, Other:$false
        $colA.TextToColumns($rawSheet.Range("A1"), 1, 1, $false, $false, $isSemi, (-not $isSemi), $false, $false)
    }

    # Tentukan baris dan kolom riil
    $lastRow = $rawSheet.Cells($rawSheet.Rows.Count, "A").End(-4162).Row # -4162 = xlUp
    $lastCol = $rawSheet.Cells(1, $rawSheet.Columns.Count).End(-4159).Column # -4159 = xlToLeft
    if ($lastRow -lt 2) {
        $lastRow = $rawSheet.UsedRange.Rows.Count
    }
    if ($lastCol -lt 2) {
        $lastCol = $rawSheet.UsedRange.Columns.Count
    }
    $dataRange = $rawSheet.Range("A1", $rawSheet.Cells.Item($lastRow, $lastCol))

    # 3. Buat Sheet Baru untuk Pivot Table
    $pivotSheet = $wb.Sheets.Add()
    $pivotSheet.Name = "Sheet1"

    $pivotCache = $wb.PivotCaches().Create(1, $dataRange) # 1 = xlDatabase
    $pivotTable = $pivotCache.CreatePivotTable($pivotSheet.Range("A3"), "PivotDO")

    # 4. Report layout: Show in Tabular Form (1 = xlTabularRow)
    $pivotTable.RowAxisLayout(1)

    # Tambahkan Row Fields: WILMGR, WILASM, WILGL
    $fMgr = $pivotTable.PivotFields("WILMGR")
    $fMgr.Orientation = 1 # xlRowField
    $fMgr.Position = 1

    $fAsm = $pivotTable.PivotFields("WILASM")
    $fAsm.Orientation = 1 # xlRowField
    $fAsm.Position = 2

    $fGl = $pivotTable.PivotFields("WILGL")
    $fGl.Orientation = 1 # xlRowField
    $fGl.Position = 3

    # Tambahkan Data Fields: Sum of qty, Sum of rp (Format General murni tanpa koma desimal)
    $fQty = $pivotTable.AddDataField($pivotTable.PivotFields("qty"), "Sum of qty", -4157) # -4157 = xlSum
    $fQty.NumberFormat = "General"

    $fRp = $pivotTable.AddDataField($pivotTable.PivotFields("rp"), "Sum of rp", -4157)
    $fRp.NumberFormat = "General"

    # Tambahkan Page Field: mstatus
    $fStatus = $pivotTable.PivotFields("mstatus")
    $fStatus.Orientation = 3 # xlPageField
    $fStatus.Position = 1
    $fStatus.EnableMultiplePageItems = $true

    # Filter out Cancel dan Blanks, pertahankan Open dan Close
    foreach ($item in $fStatus.PivotItems()) {
        if ($item.Name -like "*Cancel*" -or $item.Name -eq "(blank)" -or $item.Name -like "*blank*") {
            $item.Visible = $false
        } else {
            $item.Visible = $true
        }
    }

    # Sembunyikan blank pada Row Fields jika ada
    foreach ($fld in @($fMgr, $fAsm, $fGl)) {
        foreach ($item in $fld.PivotItems()) {
            if ($item.Name -eq "(blank)" -or $item.Name -like "*blank*") {
                try { $item.Visible = $false } catch {}
            }
        }
    }

    # 5. Kunci Urutan Daerah Sesuai Standar Operasional Cabang Makassar
    # Urutan Manager
    $mgrPatterns = @("*MANAGER 1*", "*MANAGER 2*")
    $pos = 1
    foreach ($pat in $mgrPatterns) {
        foreach ($it in $fMgr.PivotItems()) {
            if ($it.Name -like $pat) {
                try { $it.Position = $pos; $pos++ } catch {}
            }
        }
    }

    # Urutan ASM
    $asmPatterns = @(
        "*ASM 1 MAKASSAR*",
        "*ASM PARE-PARE*",
        "*ASM KENDARI*",
        "*ASM 2 MAKASSAR*",
        "*ASM BONE*",
        "*ASM PALOPO*",
        "*GL NB MAKASSAR*"
    )
    $pos = 1
    foreach ($pat in $asmPatterns) {
        foreach ($it in $fAsm.PivotItems()) {
            if ($it.Name -like $pat) {
                try { $it.Position = $pos; $pos++ } catch {}
            }
        }
    }

    # Urutan GL / Unit
    $glPatterns = @(
        # Under ASM 1 MAKASSAR
        "*GL SD 1 MAKASSAR*",
        "*GL SD 2 MAKASSAR*",
        "*REP PANGKEP*",
        # Under ASM PARE-PARE
        "*GL PARE-PARE*",
        "*REP SIDRAP*",
        "*REP MAMUJU*",
        # Under ASM KENDARI (ASM KENDARI di awal / beginning)
        "*ASM KENDARI*",
        "*GL TK*KENDARI*",
        "*GL SMP*KENDARI*",
        "*REP KOLAKA*",
        "*REP BUTON*",
        # Under ASM 2 MAKASSAR
        "*GL SMP MAKASSAR*",
        "*GL SMA*MAKASSAR*",
        # Under ASM BONE
        "*REP BULUKUMBA*",
        "*GL BONE*",
        "*REP BANTAENG*",
        # Under ASM PALOPO
        "*GL PALOPO*",
        "*REP MASAMBA*",
        "*REP TORAJA*",
        # Under GL NB MAKASSAR
        "*GL NB MAKASSAR*"
    )
    $pos = 1
    foreach ($pat in $glPatterns) {
        foreach ($it in $fGl.PivotItems()) {
            if ($it.Name -like $pat) {
                try { $it.Position = $pos; $pos++ } catch {}
            }
        }
    }

    # 6. Pastikan seluruh tipe data menggunakan General (tanpa koma desimal di belakang)
    $rawSheet.Cells.NumberFormat = "General"
    $pivotSheet.Columns.AutoFit() | Out-Null
    $rawSheet.Columns.AutoFit() | Out-Null

    # 7. Simpan file sebagai XLSX (51 = xlOpenXMLWorkbook)
    $saveTarget = $outXlsx
    try {
        if (Test-Path $saveTarget) {
            Remove-Item $saveTarget -Force -ErrorAction Stop
        }
    }
    catch {
        # Jika file sedang dibuka oleh Excel/aplikasi lain, simpan dengan akhiran _baru
        $baseDir = Split-Path $outXlsx -Parent
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($outXlsx)
        $ext = [System.IO.Path]::GetExtension($outXlsx)
        $saveTarget = Join-Path $baseDir "${baseName}_baru${ext}"
        if (Test-Path $saveTarget) { Remove-Item $saveTarget -Force -ErrorAction SilentlyContinue }
    }
    $wb.SaveAs($saveTarget, 51)
    $wb.Close($false)
    Write-Output "SUCCESS: $saveTarget"
}
catch {
    Write-Error "ERROR: $_"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
'''

def parse_target_date(date_str_or_dir: Optional[str]) -> Optional[Tuple[int, int, int]]:
    """
    Mengekstrak (year, month, day) dari string tanggal atau nama folder.
    Contoh: '13 OKT 2026' -> (2026, 10, 13)
    """
    if not date_str_or_dir:
        return None
    base = os.path.basename(str(date_str_or_dir)).upper()
    month_map = {
        'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MEI': 5, 'MAY': 5,
        'JUN': 6, 'JUL': 7, 'AGU': 8, 'AUG': 8, 'SEP': 9, 'OKT': 10,
        'OCT': 10, 'NOV': 11, 'DES': 12, 'DEC': 12
    }
    m = re.search(r'(\d{1,2})\s*([A-Z]{3})\s*(\d{4})', base)
    if m:
        d = int(m.group(1))
        mon = month_map.get(m.group(2)[:3], 0)
        y = int(m.group(3))
        if mon > 0:
            return y, mon, d
    m2 = re.search(r'(\d{4})[-_](\d{1,2})[-_](\d{1,2})', base)
    if m2:
        return int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
    return None

def is_do_summary_file(filepath: str) -> bool:
    """
    Memvalidasi apakah file CSV merupakan file DO Summary dari Axapta:
    Mengecek header dan kata kunci seperti DELIVERYORDERID, 56-DOE-, DOWNLOADTIME, WILMGR.
    """
    if not os.path.isfile(filepath):
        return False
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as fp:
            for _ in range(35):
                line = fp.readline()
                if not line:
                    break
                up = line.upper()
                if 'DELIVERYORDERID' in up or '56-DOE' in up or ('WILMGR' in up and 'WILASM' in up) or 'DOWNLOADTIME' in up:
                    return True
    except Exception:
        pass
    return False

def detect_csv_delimiter(csv_path: str) -> str:
    """Deteksi otomatis delimiter CSV (koma atau titik koma)."""
    try:
        with open(csv_path, 'r', encoding='utf-8', errors='ignore') as fp:
            for _ in range(5):
                line = fp.readline()
                if ';' in line and line.count(';') > line.count(','):
                    return ';'
                if ',' in line and line.count(',') >= line.count(';'):
                    return ','
    except Exception:
        pass
    return ','

def evaluate_do_file_match(filepath: str, y: int, mon: int, d: int) -> Tuple[int, str]:
    """
    Mengevaluasi kesesuaian file DO terhadap tanggal target (y, mon, d):
    - Skor 100: Tanggal transaksi di dalam isi file (kolom tgl, tgl_jam, DOWNLOADTIME) persis cocok.
    - Skor 80: Token tanggal Axapta pada nama file (misal 20261013SD20261013) cocok.
    - Skor 60: File modification time (mtime) cocok dengan hari tersebut.
    """
    month_abbrs = ['', 'JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    target_mon_en = month_abbrs[mon] if 1 <= mon <= 12 else ''
    target_mon_id = {'MAY': 'MEI', 'AUG': 'AGU', 'OCT': 'OKT', 'DEC': 'DES'}.get(target_mon_en, target_mon_en)
    y2 = str(y)[2:]
    
    score = 0
    reasons = []

    # 1. Pengecekan isi data (Bobot Utama)
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as fp:
            for _ in range(60):
                line = fp.readline()
                if not line:
                    break
                # Format kolom tgl: '13 Oct 26' / '06 Oct 26' / '13 Okt 26'
                pat_tgl = rf'\b0?{d}\s+(?:{target_mon_en}|{target_mon_id})\s+{y2}\b'
                if re.search(pat_tgl, line, re.IGNORECASE):
                    score = max(score, 100)
                    reasons.append(f"Data transaksi tanggal '{d} {target_mon_en} {y2}'")
                    break
                # Format DOWNLOADTIME: YYMMDD -> '261013:'
                dl_pat = rf'{y2}{mon:02d}0?{d}:'
                if re.search(dl_pat, line):
                    score = max(score, 100)
                    reasons.append(f"Header DOWNLOADTIME '{y2}{mon:02d}{d:02d}'")
                    break
                # Format tgl_jam: 'Oct 13 2026' / 'Oct  6 2026'
                pat_jam = rf'\b(?:{target_mon_en}|{target_mon_id})\s+0?{d}\s+{y}\b'
                if re.search(pat_jam, line, re.IGNORECASE):
                    score = max(score, 100)
                    reasons.append(f"Data tgl_jam '{target_mon_en} {d} {y}'")
                    break
    except Exception:
        pass

    # 2. Pengecekan token Axapta pada nama file
    fname = os.path.basename(filepath).upper()
    toks = [
        f"{y}{mon}{d}SD",
        f"{y}{mon:02d}{d:02d}SD",
        f"{y}{mon}{d:02d}SD",
        f"{y}{mon:02d}{d}SD"
    ]
    for tok in toks:
        if tok in fname:
            if score < 80:
                score = 80
                reasons.append(f"Token Axapta '{tok}' di nama file")
            break

    pat_fn = rf'\b0?{d}[\s_.-]*(?:{target_mon_en}|{target_mon_id})[\s_.-]*{y}\b'
    if re.search(pat_fn, fname):
        if score < 80:
            score = 80
            reasons.append("Pola tanggal pada nama file")

    # 3. Pengecekan waktu modifikasi file (mtime)
    try:
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(filepath))
        if mtime.year == y and mtime.month == mon and mtime.day == d:
            if score < 60:
                score = 60
                reasons.append(f"Waktu download/modifikasi file ({mtime.strftime('%d-%m-%Y %H:%M')})")
    except Exception:
        pass

    return score, "; ".join(reasons)

def find_do_summary_for_date(
    work_dir: str,
    target_date_str: Optional[str] = None,
    search_dirs: Optional[List[str]] = None,
    report_fn: Callable[[str], None] = print
) -> Optional[str]:
    """
    Deteksi DO Summary Berbasis Folder (Folder-Based Detection):
    Prinsip: Pengguna selalu menaruh file DO Summary di dalam folder kerja harian (work_dir).
    Sistem langsung memprioritaskan file CSV yang ada di dalam folder tersebut tanpa
    tergantung pada nama file yang kaku/acak dari sistem Axapta.
    """
    if not work_dir or not os.path.exists(work_dir):
        return None

    # =========================================================================
    # 1. PRIORITAS UTAMA: DETEKSI LANGSUNG DI DALAM FOLDER KERJA (BERBASIS FOLDER)
    # =========================================================================
    csv_in_folder = [
        f for f in glob.glob(os.path.join(work_dir, "*.csv"))
        if not os.path.basename(f).startswith("~$") and not os.path.basename(f).startswith("_temp")
    ]

    if csv_in_folder:
        # Jika hanya ada 1 file CSV di folder kerja, langsung gunakan (karena pengguna selalu menaruhnya di sini)
        if len(csv_in_folder) == 1:
            chosen = csv_in_folder[0]
            report_fn(f"[DETEKSI DO] Ditemukan file DO Summary di folder kerja: {os.path.basename(chosen)}")
            return chosen

        # Jika ada beberapa file CSV di folder kerja:
        # Prioritas A: File dengan nama memuat kata kunci DO / SUMMARY / EXTERNAL
        keyword_matches = [
            f for f in csv_in_folder
            if any(k in os.path.basename(f).upper() for k in ["SUMMARY", "DO", "EXTERNAL", "DOE"])
        ]
        if keyword_matches:
            keyword_matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            chosen = keyword_matches[0]
            report_fn(f"[DETEKSI DO] Ditemukan file DO Summary di folder kerja: {os.path.basename(chosen)}")
            return chosen

        # Prioritas B: File yang tervalidasi struktur kolom DO Axapta
        verified_matches = [f for f in csv_in_folder if is_do_summary_file(f)]
        if verified_matches:
            verified_matches.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            chosen = verified_matches[0]
            report_fn(f"[DETEKSI DO] Ditemukan file DO Summary di folder kerja: {os.path.basename(chosen)}")
            return chosen

        # Prioritas C: File CSV terbaru di folder kerja
        csv_in_folder.sort(key=lambda x: os.path.getmtime(x), reverse=True)
        chosen = csv_in_folder[0]
        report_fn(f"[DETEKSI DO] Menggunakan file CSV di folder kerja: {os.path.basename(chosen)}")
        return chosen

    # Jika tidak ada file CSV di folder kerja
    return None

def process_do_summary_file(csv_path: str, out_xlsx_path: Optional[str] = None, log_func: Callable[[str], None] = print) -> str:
    """
    Memproses file Do_Summary_External_*.csv:
    1. Text to columns (koma atau titik koma)
    2. Sheet baru dengan Pivot Table (WILMGR -> WILASM -> WILGL)
    3. Values: Sum of qty, Sum of rp
    4. Filters: mstatus Open & Close
    5. Layout: Tabular form dengan urutan wilayah yang presisi.
    """
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"File CSV tidak ditemukan: {csv_path}")

    csv_abs = os.path.abspath(csv_path)
    if out_xlsx_path is None:
        # Contoh: Do_Summary_External_2026106SD2026106.csv -> Do_Summary_External_2026106SD2026106.csv.xlsx
        out_xlsx_path = csv_abs + ".xlsx"
    out_abs = os.path.abspath(out_xlsx_path)

    log_func(f"Memproses DO Summary: {os.path.basename(csv_abs)}")
    log_func(f"-> Target output: {os.path.basename(out_abs)}")

    # Deteksi delimiter CSV
    delim = detect_csv_delimiter(csv_abs)
    if delim == ";":
        log_func("-> Format delimiter terdeteksi: Titik Koma (;)")

    # Buat file script powershell sementara
    temp_ps1 = os.path.join(os.path.dirname(csv_abs), "_temp_do_pivot.ps1")
    with open(temp_ps1, "w", encoding="utf-8") as f:
        f.write(PS_SCRIPT_TEMPLATE)

    try:
        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", temp_ps1,
            "-csvPath", csv_abs,
            "-outXlsx", out_abs,
            "-delimiter", delim
        ]
        log_func("-> Menjalankan engine Excel untuk membuat Pivot Table Tabular...")
        res = subprocess.run(cmd, capture_output=True, text=True, check=True)
        log_func("-> " + res.stdout.strip())
        log_func(f"[OK] Berhasil membuat file: {os.path.basename(out_abs)}")
        return out_abs
    except subprocess.CalledProcessError as e:
        err_msg = e.stderr or e.stdout
        log_func(f"[ERROR] Gagal memproses DO Summary: {err_msg}")
        raise RuntimeError(f"Gagal memproses DO Summary: {err_msg}")
    finally:
        if os.path.exists(temp_ps1):
            try:
                os.remove(temp_ps1)
            except Exception:
                pass

if __name__ == "__main__":
    test_csv = r"C:\Users\Akmal\Akmal\etc\Do_Summary_External_2026106SD2026106.csv"
    if os.path.exists(test_csv):
        out = process_do_summary_file(test_csv)
        print("Done:", out)
