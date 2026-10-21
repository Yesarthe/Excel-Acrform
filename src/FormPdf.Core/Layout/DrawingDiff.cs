using System;
using System.Collections.Generic;
using FormPdf.Core.Model;

namespace FormPdf.Core.Layout
{
    /// <summary>
    /// 2 つの PDF の描画の差分。
    /// 「チェックボックス・オプションボタンを表示して出力した PDF」から
    /// 「非表示にして出力した PDF」の描画を取り除くと、コントロール自身が描いたもの（四角・丸）だけが残る。
    /// 罫線・網掛け・他の図形は両方に同じ座標で現れるので消える。
    /// </summary>
    public static class DrawingDiff
    {
        /// <summary>同じ描画とみなす座標の誤差（pt）。同じ Excel で出力した PDF なので座標は実質一致する。</summary>
        public const double Tolerance = 0.05;

        /// <returns>withItems にあって withoutItems に無い描画（同じものが複数あれば個数で相殺する）。</returns>
        public static List<FilledArea> Subtract(IEnumerable<FilledArea> withItems, IEnumerable<FilledArea> withoutItems)
        {
            // ページと左下座標（1pt 単位）でバケットに分けて照合する
            var buckets = new Dictionary<(int page, long x, long y), List<PdfBox>>();
            foreach (var a in withoutItems)
            {
                var key = Key(a.PageIndex, a.Bounds);
                if (!buckets.TryGetValue(key, out var list)) buckets[key] = list = new List<PdfBox>();
                list.Add(a.Bounds);
            }

            var result = new List<FilledArea>();
            foreach (var a in withItems)
            {
                if (!RemoveMatch(buckets, a.PageIndex, a.Bounds))
                    result.Add(a);
            }
            return result;
        }

        private static (int, long, long) Key(int page, PdfBox b) =>
            (page, (long)Math.Floor(b.X), (long)Math.Floor(b.Y));

        private static bool RemoveMatch(Dictionary<(int, long, long), List<PdfBox>> buckets, int page, PdfBox b)
        {
            var (_, kx, ky) = Key(page, b);
            // 誤差で隣のバケットに入っている場合も探す
            for (long dx = -1; dx <= 1; dx++)
                for (long dy = -1; dy <= 1; dy++)
                {
                    if (!buckets.TryGetValue((page, kx + dx, ky + dy), out var list)) continue;
                    for (int i = 0; i < list.Count; i++)
                    {
                        var o = list[i];
                        if (Math.Abs(o.X - b.X) <= Tolerance && Math.Abs(o.Y - b.Y) <= Tolerance &&
                            Math.Abs(o.Width - b.Width) <= Tolerance && Math.Abs(o.Height - b.Height) <= Tolerance)
                        {
                            list.RemoveAt(i);
                            return true;
                        }
                    }
                }
            return false;
        }
    }
}
