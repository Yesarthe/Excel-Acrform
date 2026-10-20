using System;
using System.IO;
using System.Linq;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;
using FormPdf.Core.Pdf;
using iText.Forms;
using iText.Forms.Fields;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using Xunit;

namespace FormPdf.Core.Tests
{
    public class WriterTests : IDisposable
    {
        private readonly string _dir;

        public WriterTests()
        {
            _dir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "formpdf-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_dir);
            Logger.LogPathOverride = System.IO.Path.Combine(_dir, "log.txt");
        }

        public void Dispose()
        {
            try { Directory.Delete(_dir, true); } catch { }
        }

        /// <summary>Linux のテスト環境にある日本語フォント（IPA ゴシック）を使う。Windows では MS ゴシック等。</summary>
        private static FormPdfSettings TestSettings()
        {
            var s = new FormPdfSettings();
            if (!OperatingSystem.IsWindows())
            {
                s.FontName = "IPAGothic";
                s.FallbackFonts.Clear();
            }
            return s;
        }

        private static FontLocator TestLocator() => new FontLocator(new[] { "/usr/share/fonts" });

        private string MakeSource(int pages)
        {
            var path = System.IO.Path.Combine(_dir, "src.pdf");
            using (var pdf = new PdfDocument(new PdfWriter(path)))
            {
                for (int i = 0; i < pages; i++)
                {
                    var page = pdf.AddNewPage(PageSize.A4);
                    new PdfCanvas(page).Rectangle(50, 50, 100, 100).Stroke();
                }
            }
            return path;
        }

        private static FormDocumentSpec SampleSpec(string parent = null)
        {
            var spec = new FormDocumentSpec();
            spec.Fields.Add(new TextFieldSpec
            {
                ParentName = parent, Name = "氏名", PageIndex = 0, Rect = new PdfBox(100, 700, 200, 20),
                Value = "日本語テスト①", FontSize = 10.5, Alignment = HAlign.Left,
            });
            spec.Fields.Add(new TextFieldSpec
            {
                ParentName = parent, Name = "備考", PageIndex = 0, Rect = new PdfBox(100, 600, 200, 60),
                Value = "1行目\n2行目", FontSize = 10, Multiline = true,
            });
            spec.Fields.Add(new TextFieldSpec
            {
                ParentName = parent, Name = "金額", PageIndex = 1, Rect = new PdfBox(100, 500, 100, 20),
                Value = "12,345", FontSize = 11, Alignment = HAlign.Right,
            });
            spec.Fields.Add(new CheckBoxSpec { ParentName = parent, Name = "同意", PageIndex = 0, Rect = new PdfBox(100, 550, 10, 10), Checked = true });
            spec.Fields.Add(new CheckBoxSpec { ParentName = parent, Name = "確認", PageIndex = 0, Rect = new PdfBox(100, 530, 10, 10), Checked = false });
            var g = new RadioGroupSpec { ParentName = parent, Name = "性別", SelectedValue = "Option Button 2" };
            g.Buttons.Add(new RadioButtonSpec { ExportValue = "Option Button 1", PageIndex = 0, Rect = new PdfBox(100, 480, 10, 10) });
            g.Buttons.Add(new RadioButtonSpec { ExportValue = "Option Button 2", PageIndex = 0, Rect = new PdfBox(160, 480, 10, 10) });
            g.Buttons.Add(new RadioButtonSpec { ExportValue = "Option Button 3", PageIndex = 1, Rect = new PdfBox(220, 480, 10, 10) });
            spec.Fields.Add(g);
            var jp = new RadioGroupSpec { ParentName = parent, Name = "選択", SelectedValue = "いいえ" };
            jp.Buttons.Add(new RadioButtonSpec { ExportValue = "はい", PageIndex = 1, Rect = new PdfBox(100, 400, 10, 10) });
            jp.Buttons.Add(new RadioButtonSpec { ExportValue = "いいえ", PageIndex = 1, Rect = new PdfBox(160, 400, 10, 10) });
            spec.Fields.Add(jp);
            return spec;
        }

        [Fact]
        public void WritesFields_WithEmbeddedJapaneseFont_AndAppearances()
        {
            var src = MakeSource(2);
            var dst = System.IO.Path.Combine(_dir, "out.pdf");
            var warnings = new WarningList();
            var writer = new ITextPdfFormWriter(TestLocator());

            int n = writer.Write(src, dst, SampleSpec(), TestSettings(), warnings);
            Assert.Empty(warnings.Items);
            Assert.Equal(10, n);
            CopyForInspection(dst, "out.pdf");

            using var pdf = new PdfDocument(new PdfReader(dst));
            var form = PdfFormCreator.GetAcroForm(pdf, false);
            var fields = form.GetAllFormFields();

            Assert.Equal("日本語テスト①", fields["氏名"].GetValueAsString());
            Assert.Equal("Yes", fields["同意"].GetValueAsString());
            Assert.Equal("Off", fields["確認"].GetValueAsString());
            Assert.Equal("Option Button 2", fields["性別"].GetValueAsString());

            // NeedAppearances に頼らない
            var acro = pdf.GetCatalog().GetPdfObject().GetAsDictionary(PdfName.AcroForm);
            var need = acro.GetAsBoolean(PdfName.NeedAppearances);
            Assert.True(need == null || !need.GetValue());

            // /DR にフォントが登録され、/DA がそれを参照する
            var drFonts = acro.GetAsDictionary(PdfName.DR).GetAsDictionary(PdfName.Font);
            var da = fields["氏名"].GetPdfObject().GetAsString(PdfName.DA)?.ToUnicodeString()
                     ?? acro.GetAsString(PdfName.DA).ToUnicodeString();
            var daFont = da.Split(' ').First(t => t.StartsWith("/")).Substring(1);
            Assert.True(drFonts.ContainsKey(new PdfName(daFont)), $"DA={da} DR={string.Join(",", drFonts.KeySet())}");

            // フル埋め込み（サブセット名 "ABCDEF+" が付かない）
            var type0 = drFonts.GetAsDictionary(new PdfName(daFont));
            Assert.Equal(PdfName.Type0, type0.GetAsName(PdfName.Subtype));
            var baseFont = type0.GetAsName(PdfName.BaseFont).GetValue();
            Assert.DoesNotMatch("^[A-Z]{6}\\+", baseFont);
            var descendant = type0.GetAsArray(PdfName.DescendantFonts).GetAsDictionary(0);
            var fontFile = descendant.GetAsDictionary(PdfName.FontDescriptor).GetAsStream(PdfName.FontFile2);
            Assert.NotNull(fontFile);
            Assert.True(fontFile.GetBytes().Length > 1_000_000, "フォントが丸ごと埋め込まれていない");

            // 外観ストリームがある
            foreach (var name in new[] { "氏名", "備考", "金額", "同意" })
            {
                var w = fields[name].GetWidgets()[0];
                Assert.NotNull(w.GetAppearanceDictionary());
            }

            // 右寄せ・複数行
            Assert.Equal(2, fields["金額"].GetPdfObject().GetAsNumber(PdfName.Q).IntValue());
            var ff = fields["備考"].GetPdfObject().GetAsNumber(PdfName.Ff).IntValue();
            Assert.True((ff & 4096) != 0);

            // チェックボックス: ✓（ZapfDingbats '4'）、赤
            var cbWidget = fields["同意"].GetWidgets()[0].GetPdfObject();
            Assert.Equal("4", cbWidget.GetAsDictionary(PdfName.MK).GetAsString(PdfName.CA).ToUnicodeString());
            var cbDa = (fields["同意"].GetPdfObject().GetAsString(PdfName.DA) ?? cbWidget.GetAsString(PdfName.DA)).ToUnicodeString();
            Assert.Contains("1 0 0 rg", cbDa);
            var onAp = cbWidget.GetAsDictionary(PdfName.AP).GetAsDictionary(PdfName.N).GetAsStream(new PdfName("Yes"));
            Assert.Contains("1 0 0 RG", System.Text.Encoding.ASCII.GetString(onAp.GetBytes()));
            Assert.Equal("Yes", cbWidget.GetAsName(PdfName.AS).GetValue());
            Assert.Equal("Off", fields["確認"].GetWidgets()[0].GetPdfObject().GetAsName(PdfName.AS).GetValue());

            // ページ配置
            // 日本語の Export Value は /Opt で保持
            var jp = fields["選択"].GetPdfObject();
            Assert.Equal("はい", jp.GetAsArray(PdfName.Opt).GetAsString(0).ToUnicodeString());
            Assert.Equal("いいえ", jp.GetAsArray(PdfName.Opt).GetAsString(1).ToUnicodeString());
            Assert.Equal("1", jp.GetAsName(PdfName.V).GetValue());

            Assert.Equal(1, pdf.GetPageNumber(fields["氏名"].GetWidgets()[0].GetPage()));
            Assert.Equal(2, pdf.GetPageNumber(fields["金額"].GetWidgets()[0].GetPage()));
            var radioWidgets = fields["性別"].GetWidgets();
            Assert.Equal(3, radioWidgets.Count());
            Assert.Equal(2, pdf.GetPageNumber(radioWidgets[2].GetPage()));
            Assert.Equal(4 + 2, pdf.GetPage(1).GetAnnotations().Count()); // text2 + check2 + radio2
            Assert.Equal(4, pdf.GetPage(2).GetAnnotations().Count());

            // 枠線・背景なし
            foreach (var w in fields["氏名"].GetWidgets())
            {
                var mk = w.GetPdfObject().GetAsDictionary(PdfName.MK);
                Assert.True(mk == null || (mk.Get(PdfName.BG) == null && mk.Get(PdfName.BC) == null));
            }
        }

        [Fact]
        public void HierarchicalNames_ForMultipleSheets()
        {
            var src = MakeSource(2);
            var dst = System.IO.Path.Combine(_dir, "out2.pdf");
            var spec = SampleSpec("Sheet1");
            var other = SampleSpec("Sheet2");
            spec.Fields.AddRange(other.Fields);
            var warnings = new WarningList();
            new ITextPdfFormWriter(TestLocator()).Write(src, dst, spec, TestSettings(), warnings);
            Assert.Empty(warnings.Items);
            CopyForInspection(dst, "out-multi.pdf");

            using var pdf = new PdfDocument(new PdfReader(dst));
            var fields = PdfFormCreator.GetAcroForm(pdf, false).GetAllFormFields();
            Assert.True(fields.ContainsKey("Sheet1.氏名"));
            Assert.True(fields.ContainsKey("Sheet2.氏名"));
            Assert.Equal("Option Button 2", fields["Sheet2.性別"].GetValueAsString());
            Assert.Equal(12, pdf.GetPage(1).GetAnnotations().Count());
        }

        [Fact]
        public void CalibrationMode_DrawsOverlay()
        {
            var src = MakeSource(2);
            var dst = System.IO.Path.Combine(_dir, "out3.pdf");
            var settings = TestSettings();
            settings.CalibrationMode = true;
            var warnings = new WarningList();
            new ITextPdfFormWriter(TestLocator()).Write(src, dst, SampleSpec(), settings, warnings);
            Assert.Empty(warnings.Items);
            CopyForInspection(dst, "out-calibration.pdf");
            using var pdf = new PdfDocument(new PdfReader(dst));
            Assert.True(pdf.GetPage(1).GetContentStreamCount() >= 2, "overlay stream");
        }

        [Fact]
        public void RadioGroup_WithoutSelection_IsOff_AndNoToggleToOff()
        {
            var src = MakeSource(1);
            var dst = System.IO.Path.Combine(_dir, "out5.pdf");
            var spec = new FormDocumentSpec();
            var g = new RadioGroupSpec { Name = "S_Default" };
            g.Buttons.Add(new RadioButtonSpec { ExportValue = "A", Rect = new PdfBox(10, 10, 10, 10) });
            g.Buttons.Add(new RadioButtonSpec { ExportValue = "B", Rect = new PdfBox(30, 10, 10, 10) });
            spec.Fields.Add(g);
            var warnings = new WarningList();
            int n = new ITextPdfFormWriter(TestLocator()).Write(src, dst, spec, TestSettings(), warnings);
            Assert.Equal(2, n);
            Assert.Empty(warnings.Items);

            using var pdf = new PdfDocument(new PdfReader(dst));
            var field = PdfFormCreator.GetAcroForm(pdf, false).GetField("S_Default");
            var ff = field.GetPdfObject().GetAsNumber(PdfName.Ff).IntValue();
            Assert.True((ff & PdfButtonFormField.FF_RADIO) != 0);
            Assert.True((ff & PdfButtonFormField.FF_NO_TOGGLE_TO_OFF) != 0);
            Assert.All(field.GetWidgets(), w => Assert.Equal("Off", w.GetPdfObject().GetAsName(PdfName.AS).GetValue()));
        }

        [Fact]
        public void TtcFont_IsFoundByName()
        {
            var path = TestLocator().Find(OperatingSystem.IsWindows() ? "MS ゴシック" : "WenQuanYi Zen Hei");
            Assert.NotNull(path);
            Assert.Matches(@"\.ttc,\d+$", path);
        }

        [Fact]
        public void MissingFont_FallsBackWithWarning()
        {
            var src = MakeSource(1);
            var dst = System.IO.Path.Combine(_dir, "out4.pdf");
            var settings = new FormPdfSettings { FontName = "No Such Font" };
            settings.FallbackFonts.Clear();
            var spec = new FormDocumentSpec();
            spec.Fields.Add(new TextFieldSpec { Name = "a", Rect = new PdfBox(10, 10, 100, 20), FontSize = 10, Value = "あ" });
            var warnings = new WarningList();
            new ITextPdfFormWriter(new FontLocator(new[] { _dir })).Write(src, dst, spec, settings, warnings);
            Assert.NotEmpty(warnings.Items);
            using var pdf = new PdfDocument(new PdfReader(dst));
            Assert.Equal("あ", PdfFormCreator.GetAcroForm(pdf, false).GetField("a").GetValueAsString());
        }

        private static void CopyForInspection(string path, string name)
        {
            var dir = Environment.GetEnvironmentVariable("FORMPDF_TEST_OUT");
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
                File.Copy(path, System.IO.Path.Combine(dir, name), true);
            }
        }
    }
}
