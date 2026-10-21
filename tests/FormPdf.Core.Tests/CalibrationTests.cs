using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Layout;
using FormPdf.Core.Model;
using FormPdf.Core.Pdf;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using Xunit;

namespace FormPdf.Core.Tests
{
    public class CalibrationTests : IDisposable
    {
        private readonly string _dir;

        public CalibrationTests()
        {
            _dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "formpdf-cal-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { }
        }

        private static readonly IList<PageSizeD> A4x2 = new[] { new PageSizeD(595, 842), new PageSizeD(595, 842) };

        [Fact]
        public void Palette_RoundTrips_AndAvoidsGrays()
        {
            Assert.True(MarkerCalibration.Capacity > 1000);
            for (int i = 0; i < MarkerCalibration.Capacity; i++)
            {
                var (r, g, b) = MarkerCalibration.ColorFor(i);
                Assert.False(r == g && g == b);
                Assert.Equal(i, MarkerCalibration.Decode(r, g, b));
                // PDF の色は 0〜1 の実数になるので多少の誤差を許す
                Assert.Equal(i, MarkerCalibration.Decode(r + 2.5, g - 2.5, b + 1));
                int rgb = MarkerCalibration.ExcelRgbFor(i);
                Assert.Equal((r, g, b), (rgb & 0xFF, (rgb >> 8) & 0xFF, (rgb >> 16) & 0xFF));
            }
        }

        [Theory]
        [InlineData(0, 0, 0)]
        [InlineData(255, 255, 255)]
        [InlineData(128, 128, 128)]
        [InlineData(100, 100, 100)]
        [InlineData(255, 0, 0)]
        public void CommonColors_AreNotMarkers(double r, double g, double b)
        {
            Assert.Equal(-1, MarkerCalibration.Decode(r, g, b));
        }

        private static ShapeInfo Shape(string name, double w, double h) =>
            new ShapeInfo { Name = name, Kind = TargetKind.TextBox, Bounds = new RectD(0, 0, w, h), FontSize = 11 };

        private static FilledArea Area(int marker, int page, double x, double y, double w, double h)
        {
            var (r, g, b) = MarkerCalibration.ColorFor(marker);
            return new FilledArea { PageIndex = page, R = r, G = g, B = b, Bounds = new PdfBox(x, y, w, h) };
        }

        [Fact]
        public void Apply_PicksPageWithLargerVisibleArea_AndRejectsWrongAspect()
        {
            var shapes = new List<ShapeInfo> { Shape("a", 100, 20), Shape("b", 100, 20) };
            var areas = new[]
            {
                // a: 1 ページ目は下端で大半がはみ出し、2 ページ目にほぼ全体が見える
                Area(0, 0, 50, -15, 72, 14.4),
                Area(0, 1, 50, 700, 72, 14.4),
                // b: 同じ色の正方形（網掛け等）は縦横比が違うので無視
                Area(1, 0, 10, 10, 50, 50),
                Area(1, 0, 300, 400, 72, 14.4),
            };
            int n = MarkerCalibration.Apply(areas, shapes, A4x2);
            Assert.Equal(2, n);
            Assert.Equal(1, shapes[0].MeasuredPageIndex);
            Assert.Equal(700, shapes[0].MeasuredRect.Value.Y, 3);
            Assert.Equal(0, shapes[1].MeasuredPageIndex);
            Assert.Equal(300, shapes[1].MeasuredRect.Value.X, 3);
        }

        [Fact]
        public void ReadFilledAreas_AppliesTransform_LikeExcelOutput()
        {
            // Excel の PDF のように cm で座標系を変換してから矩形を塗る
            var path = System.IO.Path.Combine(_dir, "markers.pdf");
            using (var pdf = new PdfDocument(new PdfWriter(path)))
            {
                var page = pdf.AddNewPage(PageSize.A4);
                var c = new PdfCanvas(page);
                c.SaveState()
                    .ConcatMatrix(0.72, 0, 0, -0.72, 42.5, 842 - 42.5) // 縮小 72% + 上下反転
                    .SetFillColor(ColorConstants.BLACK).Rectangle(0, 0, 500, 1).Fill(); // 罫線（無関係）
                var (r, g, b) = MarkerCalibration.ColorFor(7);
                c.SetFillColor(new DeviceRgb(r, g, b))
                    .Rectangle(100, 200, 150, 25).Fill()
                    .RestoreState();
                pdf.AddNewPage(PageSize.A4);
            }

            var areas = new ITextPdfFormWriter().ReadFilledAreas(path);
            var shapes = Enumerable.Range(0, 8).Select(i => Shape("s" + i, 150, 25)).ToList();
            Assert.Equal(1, MarkerCalibration.Apply(areas, shapes, A4x2));

            var m = shapes[7].MeasuredRect.Value;
            Assert.Equal(0, shapes[7].MeasuredPageIndex);
            Assert.Equal(42.5 + 100 * 0.72, m.X, 2);
            Assert.Equal(150 * 0.72, m.Width, 2);
            Assert.Equal(25 * 0.72, m.Height, 2);
            // 上下反転: y = 842 − 42.5 − (200 + 25) × 0.72
            Assert.Equal(842 - 42.5 - 225 * 0.72, m.Y, 2);
        }

        [Fact]
        public void Builder_UsesMeasuredPosition_AndDerivesScale()
        {
            var tb = BuilderTests.Text("氏名", 100, 100, w: 200, h: 20);
            var cb = BuilderTests.Check("同意", 100, 200);
            tb.MeasuredPageIndex = 1;
            tb.MeasuredRect = new PdfBox(60, 500, 100, 10); // 50% に縮小されている
            cb.MeasuredPageIndex = 0;
            cb.MeasuredRect = new PdfBox(80, 300, 40, 8.5);
            var sheet = BuilderTests.Sheet("S", tb, cb);
            // 計算上は 1 ページだが、実測値を使うのでページ数の警告は出ない
            var warnings = new WarningList();
            var spec = new FormSpecBuilder(new FormPdfSettings { OffsetX = 1 }, warnings).Build(new[] { sheet }, A4x2);
            Assert.Empty(warnings.Items);

            var t = spec.Fields.OfType<TextFieldSpec>().Single();
            Assert.Equal(1, t.PageIndex);
            Assert.Equal(61, t.Rect.X, 6);
            Assert.Equal(500, t.Rect.Y, 6);
            Assert.Equal(11 * 0.5, t.FontSize, 6);

            var c = spec.Fields.OfType<CheckBoxSpec>().Single();
            Assert.Equal(0, c.PageIndex);
            Assert.Equal(5, c.Rect.Width, 6); // 10pt × 0.5
            Assert.Equal(81, c.Rect.X, 6);
        }

        [Fact]
        public void Builder_FallsBackToCalculation_ForUnmeasuredShapes()
        {
            var measured = BuilderTests.Text("a", 10, 10);
            measured.MeasuredPageIndex = 0;
            measured.MeasuredRect = new PdfBox(5, 5, 100, 18);
            var calculated = BuilderTests.Text("b", 10, 40);
            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList())
                .Build(new[] { BuilderTests.Sheet("S", measured, calculated) }, A4x2);
            var b = spec.Fields.OfType<TextFieldSpec>().Single(f => f.Name == "b");
            Assert.Equal(40 + 10, b.Rect.X, 6); // 左余白 40 + Left 10
        }
    }
}
