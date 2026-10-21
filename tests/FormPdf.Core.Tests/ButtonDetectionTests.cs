using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Layout;
using FormPdf.Core.Model;
using FormPdf.Core.Pdf;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using Xunit;

namespace FormPdf.Core.Tests
{
    public class ButtonDetectionTests : IDisposable
    {
        private readonly string _dir;

        public ButtonDetectionTests()
        {
            _dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "formpdf-btn-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { }
        }

        // コントロール全体（PDF 座標）
        private static readonly PdfBox CheckControl = new PdfBox(100, 500, 80, 17);
        private static readonly PdfBox OptionControl = new PdfBox(100, 450, 80, 17);

        /// <summary>
        /// Excel の印字を模した PDF: 四角は左端から 2pt・下から 3.5pt に 10pt（立体枠＋白い中身）、
        /// キャプションは文字、セルの罫線がコントロールを横切る。丸は中心 (109, 458.5) 半径 5。
        /// </summary>
        private string MakeExcelLikePdf()
        {
            var path = System.IO.Path.Combine(_dir, "controls.pdf");
            using (var pdf = new PdfDocument(new PdfWriter(path)))
            {
                var page = pdf.AddNewPage(PageSize.A4);
                var c = new PdfCanvas(page);
                // セルの罫線（長い線）
                c.SetStrokeColor(ColorConstants.BLACK).MoveTo(40, 508).LineTo(400, 508).Stroke();
                // 四角: 外枠（灰）と中身（白）
                c.SetStrokeColor(ColorConstants.GRAY).Rectangle(102, 503.5, 10, 10).Stroke();
                c.SetFillColor(ColorConstants.WHITE).Rectangle(103, 504.5, 8, 8).Fill();
                // キャプション
                c.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA), 9)
                    .MoveText(116, 505).ShowText("Agree").EndText();
                // オプションボタンの丸
                c.SetStrokeColor(ColorConstants.GRAY).Circle(109, 458.5, 5).Stroke();
            }
            return path;
        }

        [Fact]
        public void FindsCheckBoxSquare_AndOptionCircle_FromDrawing()
        {
            var drawings = new ITextPdfFormWriter().ReadDrawnAreas(MakeExcelLikePdf()).Select(a => a.Bounds).ToList();

            var box = ButtonFrameDetector.Find(drawings, CheckControl, 10).Value;
            Assert.Equal(102, box.X, 2);
            Assert.Equal(503.5, box.Y, 2);
            Assert.Equal(10, box.Width, 2);
            Assert.Equal(10, box.Height, 2);

            var circle = ButtonFrameDetector.Find(drawings, OptionControl, 10).Value;
            Assert.Equal(104, circle.X, 1);
            Assert.Equal(453.5, circle.Y, 1);
            Assert.Equal(10, circle.Width, 1);
        }

        [Fact]
        public void ReturnsNull_WhenOnlyLinesOrLargeShapes()
        {
            var drawings = new List<PdfBox>
            {
                new PdfBox(40, 508, 360, 0),     // 罫線
                new PdfBox(100, 500, 80, 17),    // コントロール全体の背景
                new PdfBox(90, 400, 200, 200),   // グループボックス
            };
            Assert.Null(ButtonFrameDetector.Find(drawings, CheckControl, 10));
        }

        [Fact]
        public void Builder_PlacesWidgetOnDetectedSquare_ElseFallsBack()
        {
            var pdfPath = MakeExcelLikePdf();
            var cb = BuilderTests.Check("同意", 0, 0);
            cb.MeasuredPageIndex = 0;
            cb.MeasuredRect = CheckControl;
            var op = BuilderTests.Option("はい", 0, 0);
            op.MeasuredPageIndex = 0;
            op.MeasuredRect = OptionControl;
            var lone = BuilderTests.Check("印字なし", 0, 0);
            lone.MeasuredPageIndex = 0;
            lone.MeasuredRect = new PdfBox(300, 300, 80, 17);
            var sheet = BuilderTests.Sheet("S", cb, op, lone);
            var pages = new[] { new PageSizeD(595, 842) };

            var builder = new FormSpecBuilder(new FormPdfSettings(), new WarningList())
            {
                PageDrawings = new ITextPdfFormWriter().ReadDrawnAreas(pdfPath),
            };
            var spec = builder.Build(new[] { sheet }, pages);

            var checks = spec.Fields.OfType<CheckBoxSpec>().ToDictionary(f => f.Name);
            Assert.Equal(102, checks["同意"].Rect.X, 2);
            Assert.Equal(503.5, checks["同意"].Rect.Y, 2);
            // 印字が見つからないものは従来どおり左端・縦中央の 10pt
            Assert.Equal(300, checks["印字なし"].Rect.X, 2);
            Assert.Equal(303.5, checks["印字なし"].Rect.Y, 2);

            var radio = spec.Fields.OfType<RadioGroupSpec>().Single().Buttons.Single();
            Assert.Equal(104, radio.Rect.X, 1);
            Assert.Equal(453.5, radio.Rect.Y, 1);

            Assert.Equal(3, builder.ButtonCount);
            Assert.Equal(2, builder.DetectedButtonCount);

            // 設定で OFF にすると検出しない
            var off = new FormSpecBuilder(new FormPdfSettings { DetectButtonFrames = false }, new WarningList())
            {
                PageDrawings = new ITextPdfFormWriter().ReadDrawnAreas(pdfPath),
            };
            var specOff = off.Build(new[] { sheet }, pages);
            Assert.Equal(100, specOff.Fields.OfType<CheckBoxSpec>().Single(f => f.Name == "同意").Rect.X, 2);
        }
    }
}
