using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Layout;
using FormPdf.Core.Model;
using FormPdf.Core.Pdf;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using Xunit;

namespace FormPdf.Core.Tests
{
    public class DrawingDiffTests : IDisposable
    {
        private readonly string _dir;

        public DrawingDiffTests()
        {
            _dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "formpdf-diff-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { }
        }

        private static FilledArea A(int page, double x, double y, double w, double h) =>
            new FilledArea { PageIndex = page, Bounds = new PdfBox(x, y, w, h) };

        [Fact]
        public void Subtract_RemovesCommonDrawings_ByCount_WithTolerance()
        {
            var without = new[]
            {
                A(0, 10, 10, 100, 0),
                A(0, 10, 10, 100, 0),       // 同じ線が 2 本
                A(0, 9.999, 50, 5, 5),      // バケット境界をまたぐ誤差
                A(1, 10, 10, 100, 0),
            };
            var with = new[]
            {
                A(0, 10, 10, 100, 0),
                A(0, 10, 10, 100, 0),
                A(0, 10, 10, 100, 0),       // 3 本目は差分に残る
                A(0, 10.001, 50.02, 5, 5),
                A(0, 200, 300, 10, 10),     // コントロールの四角
                A(1, 10.03, 9.97, 100, 0),
                A(1, 10, 10, 100, 0),       // 別ページで 2 本目
            };
            var diff = DrawingDiff.Subtract(with, without);
            Assert.Equal(3, diff.Count);
            Assert.Contains(diff, d => d.PageIndex == 0 && d.Bounds.X == 200);
            Assert.Equal(1, diff.Count(d => d.PageIndex == 0 && d.Bounds.Width == 100));
            Assert.Equal(1, diff.Count(d => d.PageIndex == 1));
        }

        /// <summary>
        /// 紛らわしい描画がある帳票: コントロールの左端の下にセルの網掛け（小さな正方形）があり、
        /// Excel の四角は表示倍率の影響でコントロールの枠より大きく、上下にはみ出している。
        /// </summary>
        private (string with, string without) MakePdfs()
        {
            string Make(string name, bool controls)
            {
                var path = System.IO.Path.Combine(_dir, name);
                using (var pdf = new PdfDocument(new PdfWriter(path)))
                {
                    var c = new PdfCanvas(pdf.AddNewPage(PageSize.A4));
                    // 罫線とセルの網掛け（どちらの出力にもある）
                    c.SetStrokeColor(ColorConstants.BLACK).MoveTo(40, 506).LineTo(400, 506).Stroke();
                    c.SetFillColor(ColorConstants.LIGHT_GRAY).Rectangle(100.5, 500.5, 9, 9).Fill();
                    // グループボックス（どちらの出力にもある）
                    c.SetStrokeColor(ColorConstants.GRAY).Rectangle(90, 440, 200, 90).Stroke();
                    if (controls)
                    {
                        // チェックボックスの四角: 14pt、コントロール（高さ 12）より大きい
                        c.SetStrokeColor(ColorConstants.DARK_GRAY).Rectangle(103, 499, 14, 14).Stroke();
                        c.SetFillColor(ColorConstants.WHITE).Rectangle(104, 500, 12, 12).Fill();
                    }
                }
                return path;
            }
            return (Make("with.pdf", true), Make("without.pdf", false));
        }

        private static readonly PdfBox Control = new PdfBox(100, 500, 80, 12);

        [Fact]
        public void DiffIsolatesControlDrawing_AndFindsBox_WhereHeuristicFails()
        {
            var (with, without) = MakePdfs();
            var writer = new ITextPdfFormWriter();
            var all = writer.ReadDrawnAreas(with);
            var controlsOnly = DrawingDiff.Subtract(all, writer.ReadDrawnAreas(without));

            // 従来の推定（全描画から）だと網掛けを四角と誤認する
            var heuristic = ButtonFrameDetector.Find(all.Select(a => a.Bounds), Control, 10);
            Assert.True(heuristic.HasValue);
            Assert.Equal(100.5, heuristic.Value.X, 2);

            // 差分なら本物の四角だけが残る
            Assert.Equal(2, controlsOnly.Count);
            var box = ButtonFrameDetector.Find(controlsOnly.Select(a => a.Bounds), Control, 10, isolated: true).Value;
            Assert.Equal(103, box.X, 2);
            Assert.Equal(499, box.Y, 2);
            Assert.Equal(14, box.Width, 2);
        }

        [Fact]
        public void Builder_UsesIsolatedDrawings()
        {
            var (with, without) = MakePdfs();
            var writer = new ITextPdfFormWriter();
            var cb = BuilderTests.Check("同意", 0, 0);
            cb.Bounds = new RectD(0, 0, 80, 12);
            cb.MeasuredPageIndex = 0;
            cb.MeasuredRect = Control;

            var builder = new FormSpecBuilder(new FormPdfSettings(), new WarningList())
            {
                PageDrawings = DrawingDiff.Subtract(writer.ReadDrawnAreas(with), writer.ReadDrawnAreas(without)),
                PageDrawingsAreIsolated = true,
            };
            var spec = builder.Build(new[] { BuilderTests.Sheet("S", cb) }, new[] { new PageSizeD(595, 842) });
            var r = spec.Fields.OfType<CheckBoxSpec>().Single().Rect;
            Assert.Equal(103, r.X, 2);
            Assert.Equal(14, r.Width, 2);
            Assert.Equal(1, builder.DetectedButtonCount);
        }
    
        /// <summary>
        /// キャプションを消して小さくしたチェックボックス（20 x 17）。Excel は四角の他に
        /// コントロール全体の白い背景や画像を描くことがあり、それがほぼ正方形なので四角と紛れる。
        /// </summary>
        private string MakeSmallControlPdf(bool backgroundImage, bool bigLightSquare)
        {
            var path = System.IO.Path.Combine(_dir, $"small-{backgroundImage}-{bigLightSquare}.pdf");
            using (var pdf = new PdfDocument(new PdfWriter(path)))
            {
                var c = new PdfCanvas(pdf.AddNewPage(PageSize.A4));
                // 白い背景（RGB とグレーの両方）
                c.SetFillColor(new DeviceRgb(255, 255, 255)).Rectangle(100, 500, 20, 17).Fill();
                c.SetFillColor(new DeviceGray(1f)).Rectangle(100.5, 500.5, 19, 16).Fill();
                if (backgroundImage)
                {
                    var data = Enumerable.Repeat((byte)200, 4 * 4 * 3).ToArray();
                    var img = ImageDataFactory.Create(4, 4, 3, 8, data, null);
                    c.AddImageFittedIntoRectangle(img, new Rectangle(100, 500, 20, 17), false);
                }
                if (bigLightSquare)
                    c.SetFillColor(new DeviceRgb(240, 240, 240)).Rectangle(96, 494, 30, 30).Fill();
                // 本物の四角（灰色の枠と白い中身）
                c.SetStrokeColor(ColorConstants.GRAY).Rectangle(102, 503.5, 10, 10).Stroke();
                c.SetFillColor(ColorConstants.WHITE).Rectangle(103, 504.5, 8, 8).Fill();
            }
            return path;
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public void SmallControl_PicksRealBox_NotBackgroundOrLargeShapes(bool backgroundImage, bool bigLightSquare)
        {
            var path = MakeSmallControlPdf(backgroundImage, bigLightSquare);
            var drawn = new ITextPdfFormWriter().ReadDrawnAreas(path);
            Assert.Equal(3, drawn.Count(d => d.Invisible)); // 白い背景 2 個と白い中身

            var cb = BuilderTests.Check("小", 0, 0);
            cb.Bounds = new RectD(0, 0, 20, 17);
            cb.MeasuredPageIndex = 0;
            cb.MeasuredRect = new PdfBox(100, 500, 20, 17);
            var builder = new FormSpecBuilder(new FormPdfSettings(), new WarningList())
            {
                PageDrawings = drawn,           // 差分済みとみなす（すべてコントロールの描画）
                PageDrawingsAreIsolated = true,
            };
            var spec = builder.Build(new[] { BuilderTests.Sheet("S", cb) }, new[] { new PageSizeD(595, 842) });
            var r = spec.Fields.OfType<CheckBoxSpec>().Single().Rect;
            Assert.Equal(102, r.X, 2);
            Assert.Equal(503.5, r.Y, 2);
            Assert.Equal(10, r.Width, 2);
            Assert.Equal(10, r.Height, 2);
            // 周辺の描画がログに出る
            Assert.Contains(builder.ButtonLog, l => l.Contains("周辺の描画") && l.Contains("(白)"));
        }

        [Fact]
        public void OnlyWholeControlDrawing_IsUsedAsLastResort()
        {
            var control = new PdfBox(100, 500, 12, 12);
            var found = ButtonFrameDetector.Find(new[] { new PdfBox(100, 500, 12, 12) }, control, 10, isolated: true);
            Assert.Equal(12, found.Value.Width, 2);
        }
    }
}
