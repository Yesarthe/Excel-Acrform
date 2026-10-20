using System;
using System.Collections.Generic;
using System.Linq;
using FormPdf.Core.Model;

namespace FormPdf.Core.Layout
{
    /// <summary>印刷後の 1 ページ分。シート上のどの範囲が、PDF ページのどこに、何倍で載るか。</summary>
    public sealed class PageFrame
    {
        /// <summary>シート内での 0 始まりのページ番号。</summary>
        public int IndexInSheet { get; set; }
        /// <summary>このページに印刷されるシート範囲（シート座標、pt）。</summary>
        public RectD SheetRect { get; set; }
        public double Scale { get; set; }
        /// <summary>SheetRect の左上が印刷される位置（用紙左上基準、pt）。</summary>
        public double OriginX { get; set; }
        public double OriginYTop { get; set; }
        public PageSizeD PageSize { get; set; }
    }

    public sealed class SheetLayout
    {
        public List<PageFrame> Pages { get; } = new List<PageFrame>();
        public double Scale { get; set; }
    }

    /// <summary>マッピング結果（PDF 座標: 原点左下、pt）。</summary>
    public struct MappedRect
    {
        public int PageIndexInSheet;
        public PdfBox PdfRect;
        public double Scale;
    }

    /// <summary>
    /// Excel のページ設定から、図形のシート座標を印刷後のページ座標へ変換する。
    /// 改ページ位置は Excel から取得したもの（自動改ページを含む）を使い、ここでは倍率と原点だけ計算する。
    /// </summary>
    public static class PrintLayoutCalculator
    {
        private const double Eps = 0.01;

        public static SheetLayout Compute(PrintSetupInfo setup, PageSizeD pageSize, double scaleCorrectionPercent = 0)
        {
            var layout = new SheetLayout();
            if (setup.Areas == null || setup.Areas.Count == 0)
                return layout;

            double printableW = Math.Max(1, pageSize.Width - setup.LeftMargin - setup.RightMargin);
            double printableH = Math.Max(1, pageSize.Height - setup.TopMargin - setup.BottomMargin);

            double scale = ComputeScale(setup, printableW, printableH);
            scale *= 1 + scaleCorrectionPercent / 100.0;
            layout.Scale = scale;

            int index = 0;
            foreach (var area in setup.Areas)
            {
                var rowCuts = Cuts(area.Top, area.Bottom, setup.RowBreaks);
                var colCuts = Cuts(area.Left, area.Right, setup.ColumnBreaks);

                var rects = new List<RectD>();
                if (setup.Order == PageOrder.DownThenOver)
                {
                    for (int c = 0; c < colCuts.Count - 1; c++)
                        for (int r = 0; r < rowCuts.Count - 1; r++)
                            rects.Add(Cell(colCuts, rowCuts, c, r));
                }
                else
                {
                    for (int r = 0; r < rowCuts.Count - 1; r++)
                        for (int c = 0; c < colCuts.Count - 1; c++)
                            rects.Add(Cell(colCuts, rowCuts, c, r));
                }

                foreach (var rect in rects)
                {
                    double titleH = 0, titleW = 0;
                    if (setup.TitleRows.HasValue && rect.Top >= setup.TitleRows.Value.Bottom - Eps)
                        titleH = setup.TitleRows.Value.Height;
                    if (setup.TitleColumns.HasValue && rect.Left >= setup.TitleColumns.Value.Right - Eps)
                        titleW = setup.TitleColumns.Value.Width;

                    double contentW = (rect.Width + titleW) * scale;
                    double contentH = (rect.Height + titleH) * scale;

                    double ox = setup.LeftMargin + titleW * scale;
                    double oy = setup.TopMargin + titleH * scale;
                    if (setup.CenterHorizontally)
                        ox += Math.Max(0, (printableW - contentW) / 2);
                    if (setup.CenterVertically)
                        oy += Math.Max(0, (printableH - contentH) / 2);

                    layout.Pages.Add(new PageFrame
                    {
                        IndexInSheet = index++,
                        SheetRect = rect,
                        Scale = scale,
                        OriginX = ox,
                        OriginYTop = oy,
                        PageSize = pageSize,
                    });
                }
            }
            return layout;
        }

        /// <summary>
        /// Excel と同じ考え方で拡大縮小率を決める。
        /// Zoom 指定ならそのまま。FitToPages 指定なら収まる最大の整数 % （100% を超えない、10% 未満にしない）。
        /// </summary>
        public static double ComputeScale(PrintSetupInfo setup, double printableW, double printableH)
        {
            if (setup.Zoom.HasValue)
                return Clamp(setup.Zoom.Value / 100.0);

            double maxW = setup.Areas.Max(a => a.Width);
            double totalH = setup.Areas.Sum(a => a.Height);
            double maxH = setup.Areas.Max(a => a.Height);
            if (setup.TitleColumns.HasValue) maxW += setup.TitleColumns.Value.Width;
            if (setup.TitleRows.HasValue) maxH += setup.TitleRows.Value.Height;

            double s = 1.0;
            if (setup.FitToPagesWide.HasValue && setup.FitToPagesWide.Value > 0 && maxW > 0)
                s = Math.Min(s, printableW * setup.FitToPagesWide.Value / maxW);
            if (setup.FitToPagesTall.HasValue && setup.FitToPagesTall.Value > 0 && maxH > 0)
            {
                // 複数範囲はそれぞれ別ページになるので、範囲ごとに必要な高さで判定する
                double h = setup.Areas.Count > 1 ? Math.Max(maxH, totalH / setup.Areas.Count) : maxH;
                s = Math.Min(s, printableH * setup.FitToPagesTall.Value / h);
            }

            // Excel は整数 % に丸める（ページ設定画面の「拡大/縮小」に表示される値）
            s = Math.Floor(s * 100 + 1e-6) / 100.0;
            return Clamp(s);
        }

        private static double Clamp(double s) => Math.Max(0.10, Math.Min(4.00, s));

        private static List<double> Cuts(double start, double end, IEnumerable<double> breaks)
        {
            var list = new List<double> { start };
            if (breaks != null)
            {
                foreach (var b in breaks.OrderBy(x => x))
                {
                    if (b > start + Eps && b < end - Eps && b > list[list.Count - 1] + Eps)
                        list.Add(b);
                }
            }
            list.Add(end);
            return list;
        }

        private static RectD Cell(List<double> colCuts, List<double> rowCuts, int c, int r) =>
            new RectD(colCuts[c], rowCuts[r], colCuts[c + 1] - colCuts[c], rowCuts[r + 1] - rowCuts[r]);

        /// <summary>
        /// 図形の矩形をページ座標へ変換する。複数ページにまたがる場合は面積の大きいページに置く。
        /// どのページにも重ならなければ null。
        /// </summary>
        public static MappedRect? Map(SheetLayout layout, RectD shape, double offsetX = 0, double offsetY = 0)
        {
            PageFrame best = null;
            double bestArea = 0;
            foreach (var p in layout.Pages)
            {
                double a = p.SheetRect.IntersectionArea(shape);
                if (a > bestArea)
                {
                    bestArea = a;
                    best = p;
                }
            }

            // 幅や高さが 0 の図形は中心点で判定
            if (best == null && (shape.Width <= 0 || shape.Height <= 0))
                best = layout.Pages.FirstOrDefault(p => p.SheetRect.Contains(shape.CenterX, shape.CenterY));
            if (best == null)
                return null;

            double s = best.Scale;
            double x = best.OriginX + (shape.Left - best.SheetRect.Left) * s + offsetX;
            double yTop = best.OriginYTop + (shape.Top - best.SheetRect.Top) * s + offsetY;
            double w = shape.Width * s;
            double h = shape.Height * s;

            return new MappedRect
            {
                PageIndexInSheet = best.IndexInSheet,
                Scale = s,
                PdfRect = new PdfBox(x, best.PageSize.Height - yTop - h, w, h),
            };
        }
    }
}
