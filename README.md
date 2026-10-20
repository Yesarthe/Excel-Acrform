# フォームPDF出力アドイン（Excel → AcroForm PDF）

Excel のリボンのボタン一つで、帳票を **入力可能な PDF フォーム（AcroForm）** として出力するアドインです。
Excel を唯一の原本とし、改訂のたびに PDF フォームを作り直す必要をなくします。

| Excel 上のオブジェクト | PDF 上のフィールド |
|---|---|
| 図形の「テキストボックス」 | テキストフィールド（日本語入力可・フォント埋め込み） |
| フォームコントロールの「チェックボックス」 | チェックボックス（赤い ✓） |
| フォームコントロールの「オプションボタン」 | ラジオボタン（グループボックス単位で排他） |

PDF の見た目は Excel 標準の「PDF として保存」そのもの（印刷範囲・改ページ・拡大縮小・余白・ヘッダー/フッター）で、
その上にフィールドを重ねます。Acrobat Pro は不要です。

---

## 1. 要件定義からの変更点（より良い方法を採用したもの）

| 要件 | 採用した方法 | 理由 |
|---|---|---|
| VSTO アドイン + ClickOnce | **Excel-DNA アドイン（.xll）** | VSTO は Visual Studio の Office ワークロード・ClickOnce・証明書が必要で、配布と更新が重い。Excel-DNA は **.xll ファイル 1 個**（依存 DLL も同梱）を Excel に登録するだけで動き、Windows 10/11 標準の .NET Framework 4.8 で動作する。ビルドも `dotnet build` だけでよい。リボン・COM 操作・WinForms は VSTO と同じ。 |
| チェックボックス／オプションボタンも出力前に非表示 | **非表示にせず、出力の間だけ「未チェック」にする** | 非表示にすると Excel の四角・丸とキャプションが印字されず、未チェックのフィールドが PDF 上で見えなくなる（要件 4.3「キャプションは Excel の印字がそのまま残る」とも矛盾する）。未チェックで印字すれば四角・丸・キャプションは Excel のまま残り、✓ と ● だけが PDF 側で描かれるので二重表示にならない。出力後に値とリンクセルは元に戻す。テキストボックスは要件どおり非表示にする。 |
| フォントサイズは Excel の文字サイズそのまま | **印刷倍率を掛けたサイズ**（設定で OFF 可） | 「横 1 ページに収める」等で縮小印刷すると、Excel の印字は縮小されるのにフィールドの文字だけ大きくなるため。 |
| ✓ は `CheckBoxType.CHECK` | ✓ の外観を自前で描画（`/MK /CA` は ZapfDingbats の ✓、`/DA` は赤） | iText はチェックボックスの外観に色を反映しないため。どのビューアでも赤い ✓ になる。 |
| `itext7.font-asian` | 使わない（必要ない） | フォントは TrueType を Identity-H でフル埋め込みするため。非埋め込みの CJK フォントは iText のフォーム機能で例外になることも確認した。 |
| ラジオボタンの Export Value = 図形名 | 図形名のまま。ただし日本語などを含む場合は `/Opt` に格納 | PDF の名前オブジェクトは日本語を表せないため。PDF 1.7 の正規の方法（12.7.4.2.4）。 |

その他は要件定義どおりです（下記「仕様」参照）。

---

## 2. 動作環境

- Windows 10（1903 以降）/ Windows 11（.NET Framework 4.8 は標準搭載）
- Microsoft 365 Apps / Excel 2016 以降（32bit・64bit どちらも可）
- PDF の確認: Acrobat Reader / Microsoft Edge

## 3. インストール

1. 配布物（GitHub Actions の成果物 `FormPdfAddin`、またはビルドした `publish` フォルダ）を任意のフォルダに展開します。
   - `FormPdfAddin64-packed.xll` … 64bit Excel 用
   - `FormPdfAddin-packed.xll` … 32bit Excel 用
   - `Install-FormPdfAddin.ps1` / `Uninstall-FormPdfAddin.ps1`
2. **Excel を閉じて** から、PowerShell で次を実行します（管理者権限不要）。
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\Install-FormPdfAddin.ps1
   ```
   Excel のビット数を判定し、`%APPDATA%\FormPdfAddin\FormPdfAddin.xll` にコピーしてブロック解除・アドイン登録まで行います。
3. Excel を起動すると「**フォームPDF**」タブが表示されます。

<details>
<summary>手動でインストールする場合</summary>

1. `.xll` を右クリック →［プロパティ］→「セキュリティ: 許可する」にチェック（インターネットから取得したファイルは Excel がブロックするため）。
2. Excel の［ファイル］→［オプション］→［アドイン］→ 管理「Excel アドイン」［設定］→［参照］で `.xll` を選び、チェックを入れます。
   Excel のビット数は［ファイル］→［アカウント］→［Excel のバージョン情報］で確認できます。
</details>

アンインストールは `Uninstall-FormPdfAddin.ps1` を実行します。

### 署名が必要な環境（自己署名証明書での動作確認）

トラスト センターで「アプリケーション アドインに、信頼できる発行元による署名を必要とする」が有効な環境では、`.xll` に署名が必要です。
正式な証明書が無い場合の動作確認手順:

```powershell
# 1. 自己署名のコード署名証明書を作成（現在のユーザーの証明書ストア）
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=FormPdfAddin Test" -CertStoreLocation Cert:\CurrentUser\My

# 2. 「信頼されたルート証明機関」と「信頼された発行元」に登録（確認ダイアログが出ます）
Export-Certificate -Cert $cert -FilePath .\FormPdfAddinTest.cer
Import-Certificate -FilePath .\FormPdfAddinTest.cer -CertStoreLocation Cert:\CurrentUser\Root
Import-Certificate -FilePath .\FormPdfAddinTest.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher

# 3. .xll に署名
Set-AuthenticodeSignature -FilePath .\FormPdfAddin64-packed.xll -Certificate $cert -TimestampServer http://timestamp.digicert.com
```

自己署名証明書は作成した PC でしか信頼されません。他の PC に配る場合は、`.cer` を各 PC で同様にインポートするか、社内 CA の証明書を使ってください。

## 4. 使い方

1. Excel で帳票を開きます。
2. 「フォームPDF」タブの **出力範囲** を選びます（アクティブシート／選択中のシート／ブック全体）。
3. **フォームPDF出力** → 保存先を指定 → 出力後、PDF が既定のアプリで開きます。
4. 警告（フォントが無い、図形が印刷範囲外など）がある場合は最後に一覧で表示されます。

### 帳票の作り方のルール

- 入力欄にしたい場所に **テキストボックス**（［挿入］→［テキスト］→［テキストボックス］）を置きます。図形名がフィールド名になります（［数式］→［名前の管理］ではなく、図形を選択して左上の名前ボックスで変更）。
- チェックボックス・オプションボタンは **フォームコントロール**（［開発］→［挿入］→ フォームコントロール）を使います。ActiveX コントロールは対象外です。
- オプションボタンは **グループボックス** で囲むと 1 つのグループになります（グループボックス名 = PDF のフィールド名）。グループボックス外のものは「シート名_Default」グループになります。リンクセルが同じものは同じグループになります。
- 説明用の枠など、フィールドにしたくないテキストボックスは **名前を `_` で始めて** ください。そのまま印字されます。
- 非表示の図形、非表示の行・列にある図形は対象外です。グループ化した図形の中も対象になります。
- **印刷範囲を設定する** ことを推奨します（未設定の場合は使用範囲から推定するため、ずれることがあります）。

## 5. 設定

リボンの **設定** ボタンで変更できます。保存先: `%APPDATA%\FormPdfAddin\settings.json`

| 項目 | 既定 | 説明 |
|---|---|---|
| フォント名 | Yu Gothic UI | テキストフィールドのフォント。見つからなければ代替フォント（MS ゴシック → MS UI Gothic）、それも無ければ日本語を含む任意のフォント |
| フォントファイル | （空） | `.ttf/.ttc/.otf` を直接指定する場合。TTC は `C:\Windows\Fonts\msgothic.ttc,1` のように番号指定も可 |
| テキストボックスの文字を初期値にする | ON | OFF で空欄のフォームになる |
| フォントサイズを印刷倍率に合わせる | ON | OFF で Excel の文字サイズのまま |
| チェック記号の色 | #FF0000（赤） | |
| チェック／選択状態を初期値にする | ON | |
| チェックボックスの四角／オプションボタンの丸 | 10pt | Excel 上（100%）のサイズ。0 なら図形の高さの正方形。図形の左端・縦中央から切り出す |
| 四角・丸の左端位置 | 0pt | 環境によって四角の位置がずれる場合の補正 |
| 校正モード | OFF | §6 参照 |
| X / Y 補正 | 0pt | 全フィールドの位置補正（右・下が＋） |
| 倍率補正 | 0% | 印刷倍率の補正（遠くのフィールドほどずれる場合） |
| 出力後に PDF を開く | ON | |

## 6. 校正モード（位置合わせ）

1. 設定で「校正モード」を ON にして出力すると、各フィールドの位置に **半透明の赤枠とフィールド名** が描かれます。
2. Excel の印字（四角・下線）と赤枠を比べ、ずれていれば「X 補正 / Y 補正」（全体が同じ方向にずれる場合）や
   「倍率補正」（ページ右下ほどずれる場合）を調整します。
3. 合わせ終わったら校正モードを OFF に戻します。

## 7. 仕様（要約）

- **フィールド名**: 図形名。重複は `名前_2`, `名前_3`…。複数シート出力時は `シート名.図形名`（PDF の階層フィールド。シート名・図形名の `.` は `_` に置換）。
- **テキストフィールド**: フォントはサブセット化せず **フル埋め込み**（Identity-H）。`/DR` に登録し `/DA` から参照。外観ストリームを生成（`NeedAppearances` は false）。高さが文字サイズの 2 倍以上なら複数行。左/中央/右寄せを引き継ぐ。枠線・背景なし。
- **チェックボックス**: 赤い ✓。初期値は Excel の状態（リンクセルがあればその値）。ウィジェットは左端の四角部分のみ。
- **ラジオボタン**: 外観は PDF 既定（●）。Export Value = 図形名。選択済みをクリックしても解除されない（Excel と同じ）。
- **PDF 化**: Excel の `ExportAsFixedFormat`。出力の間だけテキストボックスを非表示、チェックボックス／オプションボタンを未チェックにし、**例外時も必ず元に戻す**。ブックの `Saved` フラグも元に戻すので、出力しただけでは「変更あり」になりません。
- **座標変換**: 印刷範囲・改ページ位置（Excel の `GET.DOCUMENT(64/65)`、取れなければ `HPageBreaks/VPageBreaks`）・拡大縮小率（Zoom、または FitToPages から Excel と同じく整数 % で算出）・余白・ページの方向・中央配置・印刷タイトル行/列から、図形をページ座標に変換します。ページをまたぐ図形は面積の大きいページに置きます。用紙サイズは出力された PDF から取得します。
- **ログ**: `%APPDATA%\FormPdfAddin\log.txt`（日時・ブック名・シート・件数・警告・使用フォント）。

## 8. 受け入れ確認

1. `samples\CreateSample.ps1` を実行すると、要件定義 §11 のサンプルブック `samples\sample.xlsx` が作成されます（Excel が必要なため、Excel のある PC で実行してください）。
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\samples\CreateSample.ps1
   ```
2. 確認項目:
   1. 出力 PDF の見た目が Excel の「PDF として保存」と一致する（ウィジェット以外）
   2. Acrobat Reader と Edge でテキストフィールドに「日本語テスト①」を入力・保存し、開き直しても化けない
   3. チェックボックスが赤い ✓ で表示される
   4. オプションボタンがグループ内で排他動作する（性別・連絡方法・申込書_Default）
   5. 各ウィジェットの位置ずれが ±2pt 以内（校正モードで確認）
   6. 出力後、Excel ブックに変更が残っていない（図形の表示・チェック状態・リンクセル・Saved フラグ）
   7. 「申込書」「申込書2」を 2 枚選択して出力したとき、フィールド名が `申込書.氏名` / `申込書2.氏名` となり衝突しない

## 9. ビルド

.NET SDK 8 以降があればビルドできます（Visual Studio は不要。Windows でも Linux でも可）。

```powershell
dotnet test  tests/FormPdf.Core.Tests          # 単体テスト（座標変換・フィールド組み立て・PDF 書き込み）
dotnet build src/FormPdf.AddIn -c Release       # → src/FormPdf.AddIn/bin/Release/net48/publish/*-packed.xll
```

GitHub Actions（`.github/workflows/build.yml`）でもテスト・ビルドを行い、配布物を成果物 `FormPdfAddin` としてダウンロードできます。

Visual Studio 2022 でデバッグする場合は `FormPdf.sln` を開き、`FormPdf.AddIn` をスタートアップにして実行すると Excel がアドイン付きで起動します。

### 構成

```
src/FormPdf.Core/          … COM にも Excel にも依存しない部分（単体テスト可能）
  Model/                   … 図形・ページ設定のデータ
  Layout/PrintLayoutCalculator.cs … シート座標 → PDF ページ座標
  Fields/FormSpecBuilder.cs       … フィールド名・ラジオのグループ分け・ウィジェット矩形
  Pdf/IPdfFormWriter.cs           … PDF 操作の抽象（PDFsharp 等へ差し替える場合はこれを実装）
  Pdf/ITextPdfFormWriter.cs       … iText による実装（PDF ライブラリ依存はここと FontLocator だけ）
  Infrastructure/                 … 設定（settings.json）・ログ
src/FormPdf.AddIn/         … Excel-DNA アドイン（リボン・COM 操作・設定画面）
  Excel/ShapeScanner.cs     … 図形の走査（グループ内も再帰）
  Excel/PrintSetupReader.cs … ページ設定・改ページ・ページ数の取得
  Excel/ShapeStateGuard.cs  … 出力中だけ図形を非表示／未チェックにし、必ず戻す
  ExportService.cs          … 出力処理の流れ
tests/FormPdf.Core.Tests/  … xUnit テスト
samples/CreateSample.ps1   … 受け入れ確認用サンプルブックの作成
install/                   … インストール／アンインストールスクリプト
```

## 10. ライセンス上の注意

PDF 操作に **iText（AGPL）** を使用しています。社内利用・非配布に限ってください。
ライセンス上問題がある場合は `IPdfFormWriter` を PDFsharp 等で実装し、`AddIn.cs` の `new ITextPdfFormWriter()` を差し替えてください（PDF 操作は `FormPdf.Core/Pdf` に閉じ込めてあります）。

## 11. 既知の制限・注意

- 対象外: ActiveX コントロール、ドロップダウン、リストボックス、スピンボタン、PDF 内 JavaScript、既存 PDF への追記、Mac 版・Web 版 Excel、フォントのサブセット化。
- シート保護で図形の変更が禁止されていると、出力前に図形を非表示／未チェックにできません（警告を表示し、そのまま出力します）。保護を解除してから出力してください。
- 出力中は一時的にリンクセルの値が変わるため、イベント（`Worksheet_Change` 等）は停止しています。再計算は行われます。
- フォントをフル埋め込みするため、PDF のサイズは数 MB 以上になります（フォントにより異なります）。
- Excel の列幅は印刷時にプリンタ単位へ丸められるため、横に長い帳票の右端で 1〜2pt ずれることがあります。校正モードと「倍率補正」で調整してください。
