using System;
using System.Collections.Generic;
using System.Linq;
using FormPdf.Core.Model;

namespace FormPdf.Core.Layout
{
    /// <summary>
    /// Excel が印字したチェックボックスの四角・オプションボタンの丸を PDF の描画から探す。
    /// 四角・丸の位置や大きさは Excel のバージョンや表示倍率で変わるため、決め打ちせずに実物に合わせる。
    /// </summary>
    public static class ButtonFrameDetector
    {
        private const double Tolerance = 1.0;

        /// <param name="drawings">そのページの描画（パス・画像）の外接矩形。</param>
        /// <param name="control">コントロール全体の位置（PDF 座標）。</param>
        /// <param name="expectedSize">四角・丸のおおよその大きさ（pt、印刷倍率適用後）。</param>
        /// <returns>見つかった四角・丸の外接矩形。見つからなければ null。</returns>
        public static PdfBox? Find(IEnumerable<PdfBox> drawings, PdfBox control, double expectedSize)
        {
            double minSize = Math.Max(2.5, expectedSize * 0.4);
            double maxSize = Math.Min(control.Height + Tolerance, Math.Max(expectedSize * 2.5, minSize));

            // コントロールの中にある、正方形に近い小さな図形（キャプションの文字はパスではないので含まれない）
            var squares = drawings
                .Where(b => Inside(b, control))
                .Where(b => b.Width >= minSize && b.Height >= minSize && b.Width <= maxSize && b.Height <= maxSize)
                .Where(b => b.Width / b.Height > 0.7 && b.Width / b.Height < 1.43)
                .ToList();
            if (squares.Count == 0) return null;

            // 四角・丸はコントロールの左端にある。立体枠など重なった図形はまとめて外側を取る。
            var first = squares.OrderBy(b => Math.Round(b.X, 1)).ThenByDescending(b => b.Width * b.Height).First();
            double x1 = first.X, y1 = first.Y, x2 = first.X + first.Width, y2 = first.Y + first.Height;
            foreach (var b in squares)
            {
                if (Overlaps(b, first))
                {
                    x1 = Math.Min(x1, b.X);
                    y1 = Math.Min(y1, b.Y);
                    x2 = Math.Max(x2, b.X + b.Width);
                    y2 = Math.Max(y2, b.Y + b.Height);
                }
            }
            return new PdfBox(x1, y1, x2 - x1, y2 - y1);
        }

        private static bool Inside(PdfBox b, PdfBox c) =>
            b.X >= c.X - Tolerance && b.Y >= c.Y - Tolerance &&
            b.X + b.Width <= c.X + c.Width + Tolerance && b.Y + b.Height <= c.Y + c.Height + Tolerance;

        /// <summary>b の半分以上が a と重なっている。</summary>
        private static bool Overlaps(PdfBox b, PdfBox a)
        {
            double w = Math.Min(a.X + a.Width, b.X + b.Width) - Math.Max(a.X, b.X);
            double h = Math.Min(a.Y + a.Height, b.Y + b.Height) - Math.Max(a.Y, b.Y);
            return w > 0 && h > 0 && w * h >= 0.5 * b.Width * b.Height;
        }
    }
}
