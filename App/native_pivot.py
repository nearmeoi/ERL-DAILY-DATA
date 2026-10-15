import os
import subprocess

NATIVE_PIVOT_PS_TEMPLATE = r'''
param(
    [string]$xlsxPath,
    [string]$tableName,
    [string]$targetSheet,
    [string]$rowField,
    [string]$valField,
    [string]$styleName = "PivotStyleLight16",
    [string]$pageField = "",
    [string]$pageItem = ""
)

$ErrorActionPreference = "Stop"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($xlsxPath)
    
    # Hapus sheet target lama jika sudah ada
    foreach ($sh in $wb.Sheets) {
        if ($sh.Name -eq $targetSheet) {
            $sh.Delete()
        }
    }
    
    $ws1 = $wb.Sheets.Item("Sheet1")
    $wsPivot = $wb.Sheets.Add($ws1) # Taruh di tab pertama (sebelum Sheet1)
    $wsPivot.Name = $targetSheet
    
    # Ambil range dari tabel Excel (ListObject MyGrid)
    $tbl = $ws1.ListObjects.Item($tableName)
    $dataRange = $tbl.Range
    
    $pivotCache = $wb.PivotCaches().Create(1, $dataRange) # 1 = xlDatabase
    $pt = $pivotCache.CreatePivotTable($wsPivot.Range("A3"), "PivotTable1")
    
    # Page Field (Filter Header) jika disediakan
    if ($pageField) {
        try {
            $pf = $pt.PivotFields($pageField)
            $pf.Orientation = 3 # 3 = xlPageField
            $pf.Position = 1
            if ($pageItem) {
                try { $pf.CurrentPage = $pageItem } catch {}
            }
        } catch {}
    }

    # Row Field
    $rf = $pt.PivotFields($rowField)
    $rf.Orientation = 1 # xlRowField
    $rf.Position = 1
    
    # Data Field
    $df = $pt.AddDataField($pt.PivotFields($valField), "Sum of $valField", -4157) # -4157 = xlSum
    $df.NumberFormat = "General"
    
    # Desain Style Pivot Table
    if ($styleName) {
        try { $pt.TableStyle2 = $styleName } catch {}
    }
    
    # Lebar Kolom
    $wsPivot.Columns.Item(1).ColumnWidth = 16
    $wsPivot.Columns.Item(2).ColumnWidth = 24
    
    $wb.Save()
    $wb.Close($true)
    Write-Output "SUCCESS: Native Pivot Table created with style $styleName"
}
catch {
    Write-Error "ERROR: $_"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
'''

def create_native_excel_pivot(
    xlsx_path: str,
    table_name: str = "MyGrid",
    target_sheet: str = "Sheet2",
    row_field: str = "Item number",
    val_field: str = "Quantity",
    style_name: str = "PivotStyleLight16",
    page_field: str = "",
    page_item: str = "",
    log_func=print
) -> bool:
    """
    Membuat Native Excel Pivot Table berstandar resmi Excel COM
    dengan desain PivotStyleLight16 (atau kustom), terhubung langsung ke tabel MyGrid.
    Mendukung page_field (filter, misal 'Warehouse') & page_item (default misal '56').
    """
    if not os.path.exists(xlsx_path):
        return False
        
    abs_path = os.path.abspath(xlsx_path)
    temp_ps1 = os.path.join(os.path.dirname(abs_path), "_temp_native_pivot.ps1")
    
    with open(temp_ps1, "w", encoding="utf-8") as f:
        f.write(NATIVE_PIVOT_PS_TEMPLATE)
        
    try:
        cmd = [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", temp_ps1,
            "-xlsxPath", abs_path,
            "-tableName", table_name,
            "-targetSheet", target_sheet,
            "-rowField", row_field,
            "-valField", val_field,
            "-styleName", style_name,
            "-pageField", page_field,
            "-pageItem", page_item
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, check=True)
        log_func(f"  [PIVOT DESIGN] {res.stdout.strip()}")
        return True
    except Exception as e:
        log_func(f"  [PIVOT NOTE] Tidak dapat membuat native pivot via COM: {e}")
        return False
    finally:
        if os.path.exists(temp_ps1):
            try:
                os.remove(temp_ps1)
            except Exception:
                pass
