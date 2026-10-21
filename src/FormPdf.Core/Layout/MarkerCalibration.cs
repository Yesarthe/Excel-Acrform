using System;
using System.Collections.Generic;
using FormPdf.Core.Model;

namespace FormPdf.Core.Layout
{
    /// <summary>
    /// 自動位置合わせ。
    /// 各図形と同じ位置・大きさの四角（図形ごとに固有の色）を Excel に一度 PDF 出力させ、
    /// PDF 上でその色の塗りつぶしを探して実際の位置を測る。
    /// 改ページや拡大縮小を計算で推測しないので、計算由来のずれが無くなる。
    /// </summary>
    public static class MarkerCalibration
    {
        // 1 チャンネル 12 段階（20〜240）。灰色（R=G=B）は罫線や網掛けと紛れやすいので使わない。
        private const int Levels = 12;
        private const int Step = 20;
        private const int First = 20;
        private const double Tolerance = 7;

        private static readonly List<(int R, int G, int B)> Palette = BuildPalette();
        private static readonly Dictionary<(int, int, int), int> IndexOfColor = BuildIndex();

        private static Dictionary<(int, int, int), int> BuildIndex()
        {
            var d = new Dictionary<(int, int, int), int>();
            for (int i = 0; i < Palette.Count; i++) d[Palette[i]] = i;
            return d;
        }

        private static List<(int, int, int)> BuildPalette()
        {
            var list = new List<(int, int, int)>();
            for (int r = 0; r < Levels; r++)
                for (int g = 0; g < Levels; g++)
                    for (int b = 0; b < Levels; b++)
                    {
                        if (r == g && g == b) continue;
                        list.Add((First + r * Step, First + g * Step, First + b * Step));
                    }
            // 隣り合う番号が似た色にならないよう並べ替える（決定的）
            var shuffled = new List<(int, int, int)>(list.Count);
            const int stride = 577; // list.Count と互いに素
            for (int i = 0; i < list.Count; i++)
                shuffled.Add(list[(int)((long)i * stride % list.Count)]);
            return shuffled;
        }

        /// <summary>同時に測定できる図形の数。</summary>
        public static int Capacity => Palette.Count;

        public static (int R, int G, int B) ColorFor(int index) => Palette[index];

        /// <summary>Excel の Fill.ForeColor.RGB に設定する値（R + G*256 + B*65536）。</summary>
        public static int ExcelRgbFor(int index)
        {
            var (r, g, b) = Palette[index];
            return r + (g << 8) + (b << 16);
        }

        public static int Decode(double r, double g, double b)
        {
            int ri = Level(r), gi = Level(g), bi = Level(b);
            if (ri < 0 || gi < 0 || bi < 0) return -1;
            var color = (First + ri * Step, First + gi * Step, First + bi * Step);
            return IndexOfColor.TryGetValue(color, out var index) ? index : -1;
        }

        private static int Level(double v)
        {
            double x = (v - First) / Step;
            int n = (int)Math.Round(x);
            if (n < 0 || n >= Levels) return -1;
            return Math.Abs(v - (First + n * Step)) <= Tolerance ? n : -1;
        }

        /// <summary>
        /// 検出した塗りつぶしを図形に対応づけて MeasuredPageIndex / MeasuredRect を設定する。
        /// shapes[i] がマーカー番号 i の図形。戻り値は測定できた図形の数。
        /// </summary>
        public static int Apply(IEnumerable<FilledArea> areas, IList<ShapeInfo> shapes, IList<PageSizeD> pageSizes)
        {
            var best = new Dictionary<int, (int page, PdfBox box, double visible)>();
            foreach (var a in areas)
            {
                int index = Decode(a.R, a.G, a.B);
                if (index < 0 || index >= shapes.Count) continue;
                var box = a.Bounds;
                if (!AspectMatches(shapes[index].Bounds, box)) continue;

                // ページをまたぐ図形は両方のページに描かれるので、ページ内に見えている面積の大きい方を採用
                double visible = box.Width * box.Height;
                if (a.PageIndex < pageSizes.Count)
                {
                    var p = pageSizes[a.PageIndex];
                    double w = Math.Min(box.X + box.Width, p.Width) - Math.Max(box.X, 0);
                    double h = Math.Min(box.Y + box.Height, p.Height) - Math.Max(box.Y, 0);
                    visible = w > 0 && h > 0 ? w * h : 0;
                }
                if (visible <= 0) continue;
                if (!best.TryGetValue(index, out var cur) || visible > cur.visible)
                    best[index] = (a.PageIndex, box, visible);
            }

            foreach (var kv in best)
            {
                shapes[kv.Key].MeasuredPageIndex = kv.Value.page;
                shapes[kv.Key].MeasuredRect = kv.Value.box;
            }
            return best.Count;
        }

        /// <summary>縦横比が大きく違うものは別の塗りつぶし（セルの網掛け等）とみなす。</summary>
        private static bool AspectMatches(RectD shape, PdfBox box)
        {
            if (shape.Width < 2 || shape.Height < 2 || box.Width < 1 || box.Height < 1) return true;
            double expected = shape.Width / shape.Height;
            double actual = box.Width / box.Height;
            return actual / expected > 0.75 && actual / expected < 1.33;
        }
    }
}
