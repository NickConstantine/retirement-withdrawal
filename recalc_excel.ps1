$ErrorActionPreference = "Stop"
$path = Join-Path $PSScriptRoot "Retirement Withdrawal Strategies Planner.xlsx"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.UserName = "Retirement Planner"
$wb = $xl.Workbooks.Open($path)
$xl.CalculateFullRebuild()
# strip any personal metadata so the saved file carries no author/user info
try { $wb.BuiltinDocumentProperties("Author").Value = "Retirement Planner" } catch {}
try { $wb.BuiltinDocumentProperties("Last Author").Value = "Retirement Planner" } catch {}
try { $wb.BuiltinDocumentProperties("Company").Value = "" } catch {}
try { $wb.BuiltinDocumentProperties("Manager").Value = "" } catch {}
$errors = @()
foreach ($ws in $wb.Worksheets) {
    $used = $ws.UsedRange
    foreach ($cell in $used.Cells) {
        $t = $cell.Text
        if ($t -match '^#(REF|DIV/0|VALUE|N/A|NAME|NUM|NULL)') {
            $errors += "$($ws.Name)!$($cell.Address($false,$false)) = $t"
        }
    }
}
$wb.Save()
$wb.Close($true)
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
if ($errors.Count -eq 0) { Write-Output "STATUS: success - zero formula errors" }
else { Write-Output "STATUS: errors_found"; $errors | ForEach-Object { Write-Output $_ } }
