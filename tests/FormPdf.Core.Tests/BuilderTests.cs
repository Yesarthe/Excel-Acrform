using System.Collections.Generic;
using System.Linq;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;
using Xunit;

namespace FormPdf.Core.Tests
{
    public class BuilderTests
    {
        internal static SheetInfo Sheet(string name, params ShapeInfo[] shapes)
        {
            return new SheetInfo
            {
                Name = name,
                Shapes = shapes.ToList(),
                Print = new PrintSetupInfo
                {
                    Areas = new List<RectD> { new RectD(0, 0, 500, 700) },
                    LeftMargin = 40, RightMargin = 40, TopMargin = 40, BottomMargin = 40,
                    Zoom = 100,
                },
            };
        }

        internal static ShapeInfo Text(string name, double l, double t, double w = 100, double h = 18, string text = "")
            => new ShapeInfo { Name = name, Kind = TargetKind.TextBox, Bounds = new RectD(l, t, w, h), Text = text, FontSize = 11 };

        internal static ShapeInfo Check(string name, double l, double t, bool on = false)
            => new ShapeInfo { Name = name, Kind = TargetKind.CheckBox, Bounds = new RectD(l, t, 80, 17), IsChecked = on };

        internal static ShapeInfo Option(string name, double l, double t, bool on = false, string link = null)
            => new ShapeInfo { Name = name, Kind = TargetKind.OptionButton, Bounds = new RectD(l, t, 80, 17), IsChecked = on, LinkedCell = link };

        private static readonly IList<PageSizeD> OnePage = new[] { new PageSizeD(595, 842) };

        [Fact]
        public void DuplicateNames_GetSuffix()
        {
            var sheet = Sheet("Sheet1", Text("TB", 10, 10), Text("TB", 10, 40), Check("TB", 10, 70));
            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList()).Build(new[] { sheet }, OnePage);
            Assert.Equal(new[] { "TB", "TB_2", "TB_3" }, spec.Fields.Select(f => f.FullName).ToArray());
        }

        [Fact]
        public void MultipleSheets_UseSheetPrefix_NoCollision()
        {
            var s1 = Sheet("Sheet1", Text("名前", 10, 10), Option("O1", 10, 100));
            var s2 = Sheet("Sheet2", Text("名前", 10, 10), Option("O1", 10, 100));
            var pages = new[] { new PageSizeD(595, 842), new PageSizeD(595, 842) };
            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList()).Build(new[] { s1, s2 }, pages);
            var names = spec.Fields.Select(f => f.FullName).ToList();
            Assert.Contains("Sheet1.名前", names);
            Assert.Contains("Sheet2.名前", names);
            Assert.Contains("Sheet1.Sheet1_Default", names);
            Assert.Equal(names.Count, names.Distinct().Count());
            Assert.Equal(1, spec.Fields.OfType<TextFieldSpec>().Single(f => f.ParentName == "Sheet2").PageIndex);
        }

        [Fact]
        public void SheetNameWithDot_IsSanitized()
        {
            var s1 = Sheet("v1.2", Text("A", 10, 10));
            var s2 = Sheet("Sheet2");
            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList())
                .Build(new[] { s1, s2 }, new[] { new PageSizeD(595, 842), new PageSizeD(595, 842) });
            Assert.Equal("v1_2.A", spec.Fields.Single().FullName);
        }

        [Fact]
        public void OptionButtons_GroupedByGroupBox_LinkedCell_Default()
        {
            var sheet = Sheet("S",
                Option("A1", 20, 20, on: true), Option("A2", 20, 40),          // Group Box 1 内
                Option("B1", 220, 20), Option("B2", 220, 40, on: true),         // Group Box 2 内
                Option("C1", 20, 300), Option("C2", 20, 320),                   // グループ外
                Option("L1", 20, 500, link: "S!B1"), Option("L2", 220, 500, link: "S!B1")); // リンクセル共通
            sheet.GroupBoxes.Add(new GroupBoxInfo { Name = "Group Box 1", Bounds = new RectD(10, 10, 150, 80) });
            sheet.GroupBoxes.Add(new GroupBoxInfo { Name = "Group Box 2", Bounds = new RectD(210, 10, 150, 80) });

            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList()).Build(new[] { sheet }, OnePage);
            var groups = spec.Fields.OfType<RadioGroupSpec>().ToDictionary(g => g.Name);

            Assert.Equal(new[] { "A1", "A2" }, groups["Group Box 1"].Buttons.Select(b => b.ExportValue));
            Assert.Equal("A1", groups["Group Box 1"].SelectedValue);
            Assert.Equal("B2", groups["Group Box 2"].SelectedValue);
            Assert.Equal(new[] { "C1", "C2" }, groups["S_Default"].Buttons.Select(b => b.ExportValue));
            Assert.Null(groups["S_Default"].SelectedValue);
            Assert.Equal(new[] { "L1", "L2" }, groups["S_S_B1"].Buttons.Select(b => b.ExportValue));
        }

        [Fact]
        public void LinkedCell_TakesPriorityOverGroupBox()
        {
            var sheet = Sheet("S",
                Option("A1", 20, 20, link: "S!C1"),
                Option("B1", 220, 20, link: "S!C1"));
            sheet.GroupBoxes.Add(new GroupBoxInfo { Name = "Group Box 1", Bounds = new RectD(10, 10, 150, 80) });
            sheet.GroupBoxes.Add(new GroupBoxInfo { Name = "Group Box 2", Bounds = new RectD(210, 10, 150, 80) });
            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList()).Build(new[] { sheet }, OnePage);
            var g = Assert.Single(spec.Fields.OfType<RadioGroupSpec>());
            Assert.Equal(2, g.Buttons.Count);
            Assert.Equal("Group Box 1", g.Name);
        }

        [Fact]
        public void CheckBoxWidget_IsLeftSquare()
        {
            var sheet = Sheet("S", Check("CB", 100, 100));
            var spec = new FormSpecBuilder(new FormPdfSettings { CheckBoxSize = 10 }, new WarningList()).Build(new[] { sheet }, OnePage);
            var cb = Assert.IsType<CheckBoxSpec>(spec.Fields.Single());
            Assert.Equal(140, cb.Rect.X, 3);
            Assert.Equal(10, cb.Rect.Width, 3);
            Assert.Equal(10, cb.Rect.Height, 3);
            // 縦方向は図形の中央
            double shapeBottom = 842 - (40 + 100) - 17;
            Assert.Equal(shapeBottom + 3.5, cb.Rect.Y, 3);
        }

        [Fact]
        public void TextField_MultilineAndInitialValue()
        {
            var sheet = Sheet("S", Text("one", 0, 0, h: 18, text: "abc"), Text("multi", 0, 50, h: 40, text: "a\rb\vc"));
            var spec = new FormSpecBuilder(new FormPdfSettings(), new WarningList()).Build(new[] { sheet }, OnePage);
            var t = spec.Fields.OfType<TextFieldSpec>().ToList();
            Assert.False(t[0].Multiline);
            Assert.True(t[1].Multiline);
            Assert.Equal("a\nb\nc", t[1].Value);

            var empty = new FormSpecBuilder(new FormPdfSettings { InheritTextValues = false }, new WarningList()).Build(new[] { sheet }, OnePage);
            Assert.All(empty.Fields.OfType<TextFieldSpec>(), f => Assert.Equal("", f.Value));
        }

        [Fact]
        public void ShapeOutsidePrintArea_Warns()
        {
            var warnings = new WarningList();
            var sheet = Sheet("S", Text("far", 2000, 2000));
            var spec = new FormSpecBuilder(new FormPdfSettings(), warnings).Build(new[] { sheet }, OnePage);
            Assert.Empty(spec.Fields);
            Assert.Contains(warnings.Items, w => w.Contains("far"));
        }

        [Fact]
        public void Settings_RoundTripJson()
        {
            var s = new FormPdfSettings { FontName = "MS ゴシック", OffsetX = 1.5, CalibrationMode = true, Scope = ExportScope.Workbook };
            var back = FormPdfSettings.FromJson(s.ToJson());
            Assert.Equal("MS ゴシック", back.FontName);
            Assert.Equal(1.5, back.OffsetX);
            Assert.True(back.CalibrationMode);
            Assert.Equal(ExportScope.Workbook, back.Scope);

            // 欠けた項目は既定値
            var partial = FormPdfSettings.FromJson("{\"FontName\":\"X\"}");
            Assert.Equal("X", partial.FontName);
            Assert.Equal(10, partial.CheckBoxSize);
            Assert.True(partial.InheritTextValues);
            Assert.Equal("#FF0000", partial.CheckColor);
            Assert.True(partial.AutoCalibrate); // 既存の settings.json でも自動位置合わせは ON
        }
    }
}
