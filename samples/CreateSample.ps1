<#
  受け入れ確認用のサンプルブック samples\sample.xlsx を Excel で作成する。
  （フォームコントロールは Excel でしか正しく作れないため、Excel を COM で操作して作る）

  実行: PowerShell で
    powershell -ExecutionPolicy Bypass -File .\samples\CreateSample.ps1

  内容（要件定義 §11）:
    - テキストボックス 3 個（単一行・複数行・右寄せ）
    - チェックボックス 3 個（うち 1 個は初期 ON）
    - グループボックス 2 個にオプションボタン各 3 個、グループ外オプションボタン 2 個
    - 印刷範囲 A1:H40、横 1 ページに収める縮小、上下左右余白 15mm
    - 同じ内容のシートを 2 枚（シート 2 枚選択時のフィールド名衝突の確認用）
#>
$ErrorActionPreference = 'Stop'

$path = Join-Path $PSScriptRoot 'sample.xlsx'

# Excel の定数
$msoTextOrientationHorizontal = 1
$msoAlignRight = 3
$xlCheckBox = 1
$xlGroupBox = 4
$xlOptionButton = 7
$xlOn = 1
$xlOpenXMLWorkbook = 51
$xlContinuous = 1
$xlEdgeBottom = 9

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $null
try {
    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = '申込書'

    $ws.Cells.Font.Name = 'Yu Gothic UI'
    $ws.Cells.Font.Size = 11
    # 横 1 ページに収める縮小が効くよう、A:H の合計幅を用紙幅より広くする
    $ws.Range('A:A').ColumnWidth = 14
    $ws.Range('B:H').ColumnWidth = 13
    $ws.Range('A1:H40').RowHeight = 20

    function Pos($addr) { $r = $ws.Range($addr); return @{ L = $r.Left; T = $r.Top; W = $r.Width; H = $r.Height } }

    function Add-TextBox($name, $addr, $widthCells, $heightRows, $text, [switch]$Right) {
        $p = Pos $addr
        $w = $ws.Range($addr).Resize(1, $widthCells).Width
        $h = $ws.Range($addr).Resize($heightRows, 1).Height
        $s = $ws.Shapes.AddTextbox($msoTextOrientationHorizontal, $p.L, $p.T + 1, $w, $h - 2)
        $s.Name = $name
        $s.Line.Visible = 0
        $s.Fill.Visible = 0
        $s.TextFrame2.TextRange.Text = $text
        $s.TextFrame2.TextRange.Font.Size = 11
        if ($Right) { $s.TextFrame2.TextRange.ParagraphFormat.Alignment = $msoAlignRight }
        # 下線（罫線）は Excel のセル側で描く
        $border = $ws.Range($addr).Resize($heightRows, $widthCells).Borders.Item($xlEdgeBottom)
        $border.LineStyle = $xlContinuous
    }

    function Add-Control($type, $name, $addr, $widthCells, $caption, [int]$dy = 0) {
        $p = Pos $addr
        $w = $ws.Range($addr).Resize(1, $widthCells).Width
        $s = $ws.Shapes.AddFormControl($type, $p.L + 2, $p.T + 2 + $dy, $w - 4, 17)
        $s.Name = $name
        $s.OLEFormat.Object.Caption = $caption
        return $s
    }

    $ws.Range('A1').Value2 = 'フォームPDF 受け入れ確認用サンプル'
    $ws.Range('A1').Font.Size = 16
    $ws.Range('A1').Font.Bold = $true

    # テキストボックス
    $ws.Range('A3').Value2 = '氏名'
    Add-TextBox '氏名' 'B3' 4 1 '山田 太郎'
    $ws.Range('A5').Value2 = '備考（複数行）'
    Add-TextBox '備考' 'B5' 6 4 "1 行目`r2 行目"
    $ws.Range('A10').Value2 = '金額'
    Add-TextBox '金額' 'B10' 2 1 '12,345' -Right
    $ws.Range('D10').Value2 = '円'

    # チェックボックス
    $ws.Range('A12').Value2 = '確認事項'
    $cb1 = Add-Control $xlCheckBox '同意する' 'B12' 2 '同意する'
    $cb1.ControlFormat.Value = $xlOn
    $null = Add-Control $xlCheckBox 'メール希望' 'D12' 2 'メールを希望'
    $null = Add-Control $xlCheckBox '郵送希望' 'F12' 2 '郵送を希望'

    # グループボックス 2 個 × オプションボタン 3 個
    $g1 = Pos 'A14'
    $gb1 = $ws.Shapes.AddFormControl($xlGroupBox, $g1.L + 2, $g1.T, $ws.Range('A14:C14').Width - 4, $ws.Range('A14:A18').Height)
    $gb1.Name = '性別'
    $gb1.OLEFormat.Object.Caption = '性別'
    $o = Add-Control $xlOptionButton '男性' 'A15' 2 '男性'
    $o.ControlFormat.Value = $xlOn
    $null = Add-Control $xlOptionButton '女性' 'A16' 2 '女性'
    $null = Add-Control $xlOptionButton '回答しない' 'A17' 2 '回答しない'

    $g2 = Pos 'E14'
    $gb2 = $ws.Shapes.AddFormControl($xlGroupBox, $g2.L + 2, $g2.T, $ws.Range('E14:G14').Width - 4, $ws.Range('E14:E18').Height)
    $gb2.Name = '連絡方法'
    $gb2.OLEFormat.Object.Caption = '連絡方法'
    $null = Add-Control $xlOptionButton '電話' 'E15' 2 '電話'
    $o = Add-Control $xlOptionButton 'メール' 'E16' 2 'メール'
    $o.ControlFormat.Value = $xlOn
    $null = Add-Control $xlOptionButton '郵便' 'E17' 2 '郵便'

    # グループ外のオプションボタン 2 個（→「シート名_Default」グループ）
    $ws.Range('A21').Value2 = '会員区分'
    $null = Add-Control $xlOptionButton '一般' 'B21' 2 '一般'
    $null = Add-Control $xlOptionButton '法人' 'D21' 2 '法人'

    # 変換対象外のメモ（名前が _ で始まる）
    $memo = $ws.Shapes.AddTextbox($msoTextOrientationHorizontal, (Pos 'A24').L, (Pos 'A24').T, $ws.Range('A24:H24').Width, 40)
    $memo.Name = '_説明'
    $memo.TextFrame2.TextRange.Text = 'この枠は名前が「_」で始まるため、PDF のフィールドにはならず、そのまま印字されます。'

    $ws.Range('A40').Value2 = '（印刷範囲の最終行）'

    # ページ設定: 印刷範囲 A1:H40、横 1 ページ、余白 15mm
    $ps = $ws.PageSetup
    $ps.PrintArea = '$A$1:$H$40'
    $mm15 = $xl.CentimetersToPoints(1.5)
    $ps.LeftMargin = $mm15
    $ps.RightMargin = $mm15
    $ps.TopMargin = $mm15
    $ps.BottomMargin = $mm15
    $ps.Zoom = $false
    $ps.FitToPagesWide = 1
    $ps.FitToPagesTall = $false

    # 2 枚目（同じ図形名を持つシート）
    $ws.Copy([System.Reflection.Missing]::Value, $ws)
    $wb.Worksheets.Item(2).Name = '申込書2'
    $ws.Activate()

    if (Test-Path $path) { Remove-Item $path }
    $wb.SaveAs($path, $xlOpenXMLWorkbook)
    Write-Host "作成しました: $path"
}
finally {
    if ($wb) { $wb.Close($false) }
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}
