<#
  フォームPDF出力アドインをインストールする（管理者権限は不要。現在のユーザーにだけ入る）。

  使い方: このフォルダに FormPdfAddin-packed.xll / FormPdfAddin64-packed.xll を置いて、
    powershell -ExecutionPolicy Bypass -File .\Install-FormPdfAddin.ps1
  Excel は閉じてから実行してください。
#>
$ErrorActionPreference = 'Stop'

function Get-ExcelBitness {
    # Microsoft 365 / Office 2019 以降（クイック実行版）
    foreach ($key in 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration') {
        $p = (Get-ItemProperty -Path $key -Name Platform -ErrorAction SilentlyContinue).Platform
        if ($p) { return $(if ($p -eq 'x86') { 32 } else { 64 }) }
    }
    # MSI 版 Office 2016
    foreach ($key in 'HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Outlook') {
        $b = (Get-ItemProperty -Path $key -Name Bitness -ErrorAction SilentlyContinue).Bitness
        if ($b) { return $(if ($b -eq 'x86') { 32 } else { 64 }) }
    }
    if (Test-Path 'C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE') { return 32 }
    return 64
}

if (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    throw 'Excel が起動しています。Excel をすべて閉じてから再実行してください。'
}

$bits = Get-ExcelBitness
$source = Join-Path $PSScriptRoot $(if ($bits -eq 64) { 'FormPdfAddin64-packed.xll' } else { 'FormPdfAddin-packed.xll' })
if (-not (Test-Path $source)) { throw "$source が見つかりません。" }

$dir = Join-Path $env:APPDATA 'FormPdfAddin'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$target = Join-Path $dir 'FormPdfAddin.xll'
Copy-Item $source $target -Force
# インターネットから取得したファイルは Excel にブロックされるため、ブロックを解除する
Unblock-File -Path $target

# Excel の「アドイン」に登録して有効にする
$xl = New-Object -ComObject Excel.Application
try {
    $xl.Visible = $false
    $wb = $xl.Workbooks.Add()   # ブックが無いと AddIns.Add が失敗することがある
    $addin = $xl.AddIns.Add($target, $true)
    $addin.Installed = $true
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

Write-Host "インストールしました（Excel $bits bit 用）: $target"
Write-Host 'Excel を起動すると「フォームPDF」タブが表示されます。'
