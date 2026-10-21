using System;
using System.Collections.Generic;

namespace FormPdf.Core.Model
{
    /// <summary>Excel のシート座標（左上原点・下向き正、pt）の矩形。</summary>
    public struct RectD
    {
        public double Left;
        public double Top;
        public double Width;
        public double Height;

        public RectD(double left, double top, double width, double height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        public double Right => Left + Width;
        public double Bottom => Top + Height;
        public double CenterX => Left + Width / 2;
        public double CenterY => Top + Height / 2;
        public double Area => Math.Max(0, Width) * Math.Max(0, Height);

        public bool Contains(double x, double y) =>
            x >= Left && x <= Right && y >= Top && y <= Bottom;

        public double IntersectionArea(RectD other)
        {
            double w = Math.Min(Right, other.Right) - Math.Max(Left, other.Left);
            double h = Math.Min(Bottom, other.Bottom) - Math.Max(Top, other.Top);
            return w > 0 && h > 0 ? w * h : 0;
        }

        public override string ToString() =>
            $"({Left:0.##}, {Top:0.##}, {Width:0.##} x {Height:0.##})";
    }

    /// <summary>PDF 座標（原点左下、pt）の矩形。X/Y は左下隅。</summary>
    public struct PdfBox
    {
        public double X;
        public double Y;
        public double Width;
        public double Height;

        public PdfBox(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public override string ToString() => $"[{X:0.##}, {Y:0.##}, {Width:0.##} x {Height:0.##}]";
    }

    public enum TargetKind
    {
        TextBox,
        CheckBox,
        OptionButton,
    }

    public enum HAlign
    {
        Left,
        Center,
        Right,
    }

    public enum PageOrder
    {
        /// <summary>左から右へ（Excel: 下へ、その後右へ = xlDownThenOver）</summary>
        DownThenOver,
        /// <summary>上から下へ（Excel: 右へ、その後下へ = xlOverThenDown）</summary>
        OverThenDown,
    }

    /// <summary>変換対象の図形 1 個分。COM から読み取った値をそのまま保持する。</summary>
    public sealed class ShapeInfo
    {
        public string Name { get; set; }
        public TargetKind Kind { get; set; }
        /// <summary>シート左上基準の位置（pt）。</summary>
        public RectD Bounds { get; set; }

        // テキストボックス
        public string Text { get; set; } = "";
        public double FontSize { get; set; } = 11;
        public HAlign Alignment { get; set; } = HAlign.Left;

        // チェックボックス・オプションボタン
        public bool IsChecked { get; set; }
        /// <summary>正規化済みのリンクセル（例: "Sheet1!A1"）。無ければ null。</summary>
        public string LinkedCell { get; set; }

        // 自動位置合わせ（Excel が出力した PDF 上で実測した位置）
        /// <summary>実測した PDF ページ（0 始まり、PDF 全体での番号）。未測定なら null。</summary>
        public int? MeasuredPageIndex { get; set; }
        /// <summary>実測した位置（PDF 座標）。未測定なら null。</summary>
        public PdfBox? MeasuredRect { get; set; }
    }

    /// <summary>PDF 上で塗りつぶされた領域（自動位置合わせ用のマーカー検出に使う）。</summary>
    public struct FilledArea
    {
        public int PageIndex;
        /// <summary>塗り色（0〜255）。</summary>
        public double R;
        public double G;
        public double B;
        public PdfBox Bounds;

        // 以下は ReadDrawnAreas（全描画）のときだけ設定される
        /// <summary>線で描かれている。</summary>
        public bool Stroked;
        /// <summary>塗りつぶされている。</summary>
        public bool Filled;
        /// <summary>画像。</summary>
        public bool IsImage;
        /// <summary>白一色の描画（白い紙の上では見えない。コントロールの背景など）。</summary>
        public bool Invisible;

        public string Describe()
        {
            string kind = IsImage ? "画像" : Stroked && Filled ? "線+塗" : Stroked ? "線" : "塗";
            return $"{Bounds}{kind}{(Invisible ? "(白)" : "")}";
        }
    }

    public sealed class GroupBoxInfo
    {
        public string Name { get; set; }
        public RectD Bounds { get; set; }
    }

    /// <summary>シートのページ設定のうち、座標変換に必要なものだけ。</summary>
    public sealed class PrintSetupInfo
    {
        /// <summary>印刷範囲（複数範囲可）。シート座標。</summary>
        public List<RectD> Areas { get; set; } = new List<RectD>();
        /// <summary>水平改ページ位置（改ページ直後の行の Top）。自動改ページを含む。</summary>
        public List<double> RowBreaks { get; set; } = new List<double>();
        /// <summary>垂直改ページ位置（改ページ直後の列の Left）。自動改ページを含む。</summary>
        public List<double> ColumnBreaks { get; set; } = new List<double>();

        public double LeftMargin { get; set; }
        public double RightMargin { get; set; }
        public double TopMargin { get; set; }
        public double BottomMargin { get; set; }

        /// <summary>拡大縮小率（%）。null なら FitToPages 指定。</summary>
        public double? Zoom { get; set; }
        /// <summary>横方向のページ数。null = 指定なし（自動）。</summary>
        public int? FitToPagesWide { get; set; }
        /// <summary>縦方向のページ数。null = 指定なし（自動）。</summary>
        public int? FitToPagesTall { get; set; }

        public bool CenterHorizontally { get; set; }
        public bool CenterVertically { get; set; }
        public PageOrder Order { get; set; } = PageOrder.DownThenOver;

        /// <summary>印刷タイトル行（Top/Height のみ使用）。</summary>
        public RectD? TitleRows { get; set; }
        /// <summary>印刷タイトル列（Left/Width のみ使用）。</summary>
        public RectD? TitleColumns { get; set; }
    }

    /// <summary>出力対象シート 1 枚分の情報。</summary>
    public sealed class SheetInfo
    {
        public string Name { get; set; }
        /// <summary>ワークシート以外（グラフシート等）なら false。フィールドは作らずページ数だけ数える。</summary>
        public bool IsWorksheet { get; set; } = true;
        public List<ShapeInfo> Shapes { get; set; } = new List<ShapeInfo>();
        public List<GroupBoxInfo> GroupBoxes { get; set; } = new List<GroupBoxInfo>();
        public PrintSetupInfo Print { get; set; } = new PrintSetupInfo();
        /// <summary>Excel が報告する印刷ページ数（GET.DOCUMENT(50)）。取れなければ null。</summary>
        public int? ExcelPageCount { get; set; }
    }

    /// <summary>出力 PDF のページサイズ（pt）。</summary>
    public struct PageSizeD
    {
        public double Width;
        public double Height;
        public int Rotation;

        public PageSizeD(double width, double height, int rotation = 0)
        {
            Width = width;
            Height = height;
            Rotation = rotation;
        }
    }
}
