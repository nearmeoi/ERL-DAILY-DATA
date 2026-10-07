import os
import sys
import glob
import re
import subprocess
import shutil

# Template PowerShell script untuk eksekusi Excel COM native
PS_SCRIPT_TEMPLATE = r'''
param(
    [string]$csvPath,
    [string]$outXlsx
)

$ErrorActionPreference = "Stop"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # 1. Buka file CSV
    $wb = $excel.Workbooks.Open($csvPath)
    $rawSheet = $wb.ActiveSheet

    # 2. Text to Columns pada Kolom A (delimiter koma)
    $colA = $rawSheet.Range("A:A")
    # DataType: 1=xlDelimited, ConsecutiveDelimiter:$false, Tab:$false, Semicolon:$false, Comma:$true, Space:$false, Other:$false
    $colA.TextToColumns($rawSheet.Range("A1"), 1, 1, $false, $false, $false, $true, $false, $false)

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

    # Tambahkan Data Fields: Sum of qty, Sum of rp
    $fQty = $pivotTable.AddDataField($pivotTable.PivotFields("qty"), "Sum of qty", -4157) # -4157 = xlSum
    $fQty.NumberFormat = "#,##0"

    $fRp = $pivotTable.AddDataField($pivotTable.PivotFields("rp"), "Sum of rp", -4157)
    $fRp.NumberFormat = "#,##0"

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
        # Under ASM KENDARI
        "*GL SMP/A KENDARI*",
        "*REP KOLAKA*",
        "*REP BUTON*",
        # Under ASM 2 MAKASSAR
        "*GL SMP MAKASSAR*",
        "*GL SMA/K MAKASSAR*",
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

    # 6. Auto-fit lebar kolom
    $pivotSheet.Columns.AutoFit() | Out-Null
    $rawSheet.Columns.AutoFit() | Out-Null

    # 7. Simpan file sebagai XLSX (51 = xlOpenXMLWorkbook)
    if (Test-Path $outXlsx) { Remove-Item $outXlsx -Force }
    $wb.SaveAs($outXlsx, 51)
    $wb.Close($false)
    Write-Output "SUCCESS: $outXlsx"
}
catch {
    Write-Error "ERROR: $_"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
'''

def process_do_summary_file(csv_path, out_xlsx_path=None, log_func=print):
    """
    Memproses file Do_Summary_External_*.csv:
    1. Text to columns (koma)
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
            "-outXlsx", out_abs
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
    test_csv = r"C:\Users\Akmal\Akmal\App\Do_Summary_External_2026106SD2026106.csv"
    if os.path.exists(test_csv):
        out = process_do_summary_file(test_csv)
        print("Done:", out)
