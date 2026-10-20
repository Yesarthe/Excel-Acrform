using System.Collections.Generic;
using FormPdf.Core.Layout;
using FormPdf.Core.Model;
using Xunit;

namespace FormPdf.Core.Tests
{
    public class LayoutTests
    {
        private static readonly PageSizeD A4 = new PageSizeD(595.32, 841.92);
        private const double Mm15 = 15 * 72 / 25.4; // 42.52pt

        private static PrintSetupInfo Setup(RectD area, double? zoom = 100)
        {
            return new PrintSetupInfo
            {
                Areas = new List<RectD> { area },
                LeftMargin = Mm15,
                RightMargin = Mm15,
                TopMargin = Mm15,
                BottomMargin = Mm15,
                Zoom = zoom,
            };
        }

        [Fact]
        public void Zoom100_SinglePage_MapsWithMargins()
        {
            var setup = Setup(new RectD(0, 0, 400, 600));
            var layout = PrintLayoutCalculator.Compute(setup, A4);
            Assert.Single(layout.Pages);

            var m = PrintLayoutCalculator.Map(layout, new RectD(100, 50, 80, 20)).Value;
            Assert.Equal(0, m.PageIndexInSheet);
            Assert.Equal(Mm15 + 100, m.PdfRect.X, 3);
            // y(下端) = 用紙高 − (上余白 + 50) − 20
            Assert.Equal(A4.Height - (Mm15 + 50) - 20, m.PdfRect.Y, 3);
            Assert.Equal(80, m.PdfRect.Width, 3);
            Assert.Equal(20, m.PdfRect.Height, 3);
        }

        [Fact]
        public void PrintAreaOffset_IsRelativeToAreaTopLeft()
        {
            var setup = Setup(new RectD(48, 30, 400, 600));
            var layout = PrintLayoutCalculator.Compute(setup, A4);
            var m = PrintLayoutCalculator.Map(layout, new RectD(148, 80, 10, 10)).Value;
            Assert.Equal(Mm15 + 100, m.PdfRect.X, 3);
            Assert.Equal(A4.Height - (Mm15 + 50) - 10, m.PdfRect.Y, 3);
        }

        [Fact]
        public void FitToOnePageWide_UsesIntegerPercent()
        {
            // 印刷可能幅 = 595.32 − 85.04 = 510.28。幅 700 → 72.9% → 72%
            var setup = Setup(new RectD(0, 0, 700, 600), zoom: null);
            setup.FitToPagesWide = 1;
            setup.FitToPagesTall = null;
            var layout = PrintLayoutCalculator.Compute(setup, A4);
            Assert.Equal(0.72, layout.Scale, 6);

            var m = PrintLayoutCalculator.Map(layout, new RectD(100, 100, 50, 20)).Value;
            Assert.Equal(Mm15 + 72, m.PdfRect.X, 3);
            Assert.Equal(36, m.PdfRect.Width, 3);
        }

        [Fact]
        public void FitToPages_NeverEnlarges()
        {
            var setup = Setup(new RectD(0, 0, 100, 100), zoom: null);
            setup.FitToPagesWide = 1;
            setup.FitToPagesTall = 1;
            Assert.Equal(1.0, PrintLayoutCalculator.Compute(setup, A4).Scale, 6);
        }

        [Fact]
        public void RowBreaks_SplitPages_AndShapeGoesToLargerSide()
        {
            var setup = Setup(new RectD(0, 0, 400, 1200));
            setup.RowBreaks = new List<double> { 700, 5000 /* 範囲外は無視 */ };
            var layout = PrintLayoutCalculator.Compute(setup, A4);
            Assert.Equal(2, layout.Pages.Count);

            // 690〜730: 2 ページ目側の方が大きい
            var m = PrintLayoutCalculator.Map(layout, new RectD(10, 690, 50, 40)).Value;
            Assert.Equal(1, m.PageIndexInSheet);
            Assert.Equal(A4.Height - (Mm15 + (690 - 700)) - 40, m.PdfRect.Y, 3);
        }

        [Fact]
        public void PageOrder_DownThenOver_vs_OverThenDown()
        {
            var setup = Setup(new RectD(0, 0, 1000, 1000));
            setup.RowBreaks = new List<double> { 500 };
            setup.ColumnBreaks = new List<double> { 500 };
            var shape = new RectD(600, 100, 10, 10); // 右上のページ

            setup.Order = PageOrder.DownThenOver;
            Assert.Equal(2, PrintLayoutCalculator.Map(PrintLayoutCalculator.Compute(setup, A4), shape).Value.PageIndexInSheet);

            setup.Order = PageOrder.OverThenDown;
            Assert.Equal(1, PrintLayoutCalculator.Map(PrintLayoutCalculator.Compute(setup, A4), shape).Value.PageIndexInSheet);
        }

        [Fact]
        public void CenterHorizontally_AddsOffset()
        {
            var setup = Setup(new RectD(0, 0, 400, 600));
            setup.CenterHorizontally = true;
            var layout = PrintLayoutCalculator.Compute(setup, A4);
            var m = PrintLayoutCalculator.Map(layout, new RectD(0, 0, 10, 10)).Value;
            double printable = A4.Width - 2 * Mm15;
            Assert.Equal(Mm15 + (printable - 400) / 2, m.PdfRect.X, 3);
        }

        [Fact]
        public void TitleRows_ShiftContentOnFollowingPages()
        {
            var setup = Setup(new RectD(0, 0, 400, 1200));
            setup.RowBreaks = new List<double> { 700 };
            setup.TitleRows = new RectD(0, 0, 0, 30);
            var layout = PrintLayoutCalculator.Compute(setup, A4);
            var m = PrintLayoutCalculator.Map(layout, new RectD(0, 800, 10, 10)).Value;
            Assert.Equal(1, m.PageIndexInSheet);
            Assert.Equal(A4.Height - (Mm15 + 30 + 100) - 10, m.PdfRect.Y, 3);
        }

        [Fact]
        public void OutsideArea_ReturnsNull()
        {
            var layout = PrintLayoutCalculator.Compute(Setup(new RectD(0, 0, 400, 600)), A4);
            Assert.Null(PrintLayoutCalculator.Map(layout, new RectD(500, 700, 10, 10)));
        }

        [Fact]
        public void Offsets_AreApplied()
        {
            var layout = PrintLayoutCalculator.Compute(Setup(new RectD(0, 0, 400, 600)), A4);
            var a = PrintLayoutCalculator.Map(layout, new RectD(10, 10, 10, 10)).Value;
            var b = PrintLayoutCalculator.Map(layout, new RectD(10, 10, 10, 10), 1.5, 2).Value;
            Assert.Equal(a.PdfRect.X + 1.5, b.PdfRect.X, 6);
            Assert.Equal(a.PdfRect.Y - 2, b.PdfRect.Y, 6);
        }
    }
}
