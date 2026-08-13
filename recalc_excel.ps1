$ErrorActionPreference = "Stop"
$path = Join-Path $PSScriptRoot "Retirement Withdrawal Strategies Planner.xlsx"

$xlCellTypeFormulas = -4123
$xlErrors           = 16

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.UserName = "Retirement Planner"
$wb = $xl.Workbooks.Open($path)

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$xl.CalculateFullRebuild()
Write-Output ("recalculated in {0:N1}s" -f $sw.Elapsed.TotalSeconds)

# strip any personal metadata so the saved file carries no author/user info
try { $wb.BuiltinDocumentProperties("Author").Value = "Retirement Planner" } catch {}
try { $wb.BuiltinDocumentProperties("Last Author").Value = "Retirement Planner" } catch {}
try { $wb.BuiltinDocumentProperties("Company").Value = "" } catch {}
try { $wb.BuiltinDocumentProperties("Manager").Value = "" } catch {}

# Ask Excel for error cells directly instead of walking every cell in UsedRange.
# SpecialCells throws when nothing matches, so a miss means "no errors on this sheet".
$errors = @()
foreach ($ws in $wb.Worksheets) {
    $bad = $null
    try { $bad = $ws.UsedRange.SpecialCells($xlCellTypeFormulas, $xlErrors) } catch { $bad = $null }
    if ($bad -ne $null) {
        foreach ($area in $bad.Areas) {
            foreach ($cell in $area.Cells) {
                $errors += "$($ws.Name)!$($cell.Address($false,$false)) = $($cell.Text)"
                if ($errors.Count -ge 50) { break }
            }
            if ($errors.Count -ge 50) { break }
        }
    }
    if ($errors.Count -ge 50) { break }
}

$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null

if ($errors.Count -eq 0) {
    Write-Output "STATUS: success - zero formula errors"
} else {
    Write-Output "STATUS: errors_found"
    $errors | ForEach-Object { Write-Output $_ }
}
