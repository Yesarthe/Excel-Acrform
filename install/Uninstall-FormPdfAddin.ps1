<#
  フォームPDF出力アドインをアンインストールする。Excel は閉じてから実行してください。
  設定（%APPDATA%\FormPdfAddin\settings.json）とログは残します。
#>
$ErrorActionPreference = 'Stop'

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel が起動しています。Excel をすべて閉じてから再実行してください。'
}

$target = Join-Path (Join-Path $env:APPDATA 'FormPdfAddin') 'FormPdfAddin.xll'

$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $wb = $xl.Workbooks.Add()
    foreach ($a in @($xl.AddIns)) {
        if ($a.FullName -ieq $target) { $a.Installed = $false }
    }
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

if (Test-Path $target) { Remove-Item $target -Force }
Write-Host 'アンインストールしました。'
