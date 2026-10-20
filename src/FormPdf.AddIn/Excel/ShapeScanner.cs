using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;
using Office = Microsoft.Office.Core;
using XL = Microsoft.Office.Interop.Excel;

namespace FormPdf.AddIn.Excel
{
    /// <summary>変換対象の図形と、その COM 参照（PDF 出力前に非表示・未チェックにするため保持）。</summary>
    internal sealed class TargetShape
    {
        public ShapeInfo Info { get; set; }
        public XL.Shape Com { get; set; }
    }

    /// <summary>シート上の図形を走査して、テキストボックス・チェックボックス・オプションボタン・グループボックスを集める。</summary>
    internal static class ShapeScanner
    {
        private const double XlOn = 1;

        public static void Scan(XL.Worksheet ws, SheetInfo sheet, List<TargetShape> targets, WarningList warnings)
        {
            XL.Shapes shapes = null;
            try
            {
                shapes = ws.Shapes;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    XL.Shape s = null;
                    try
                    {
                        s = shapes.Item(i);
                        if (!Visit(ws, s, sheet, targets, warnings))
                            ComScope.Release(s);
                    }
                    catch (Exception ex)
                    {
                        ComScope.Release(s);
                        warnings.Add(sheet.Name, null, $"図形 {i} を読み取れませんでした: {ex.Message}");
                    }
                }
            }
            finally
            {
                ComScope.Release(shapes);
            }
        }

        /// <returns>対象として保持した（解放してはいけない）なら true。</returns>
        private static bool Visit(XL.Worksheet ws, XL.Shape s, SheetInfo sheet, List<TargetShape> targets, WarningList warnings)
        {
            var type = s.Type;

            if (type == Office.MsoShapeType.msoGroup)
            {
                // グループ内も再帰的に（GroupItems は入れ子のグループも展開済みの末端図形を返す）
                if (IsExcluded(s)) return false;
                XL.GroupShapes items = null;
                try
                {
                    items = s.GroupItems;
                    int n = items.Count;
                    for (int j = 1; j <= n; j++)
                    {
                        XL.Shape child = null;
                        try
                        {
                            child = items.Item(j);
                            if (!Visit(ws, child, sheet, targets, warnings))
                                ComScope.Release(child);
                        }
                        catch (Exception ex)
                        {
                            ComScope.Release(child);
                            warnings.Add(sheet.Name, SafeName(s), $"グループ内の図形 {j} を読み取れませんでした: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    ComScope.Release(items);
                }
                return false;
            }

            if (type == Office.MsoShapeType.msoFormControl)
            {
                var ctl = s.FormControlType;
                if (ctl == XL.XlFormControl.xlGroupBox)
                {
                    // グループボックスは印字したまま、オプションボタンのグループ分けにだけ使う
                    sheet.GroupBoxes.Add(new GroupBoxInfo { Name = s.Name, Bounds = Bounds(s) });
                    return false;
                }
                if (ctl != XL.XlFormControl.xlCheckBox && ctl != XL.XlFormControl.xlOptionButton)
                    return false;
                if (IsExcluded(s) || IsOnHiddenCells(s)) return false;

                var info = new ShapeInfo
                {
                    Name = s.Name,
                    Kind = ctl == XL.XlFormControl.xlCheckBox ? TargetKind.CheckBox : TargetKind.OptionButton,
                    Bounds = Bounds(s),
                };
                ReadControlState(ws, s, info);
                targets.Add(new TargetShape { Info = info, Com = s });
                sheet.Shapes.Add(info);
                return true;
            }

            if (type == Office.MsoShapeType.msoTextBox)
            {
                if (IsExcluded(s) || IsOnHiddenCells(s)) return false;
                var info = new ShapeInfo
                {
                    Name = s.Name,
                    Kind = TargetKind.TextBox,
                    Bounds = Bounds(s),
                };
                ReadText(s, info);
                targets.Add(new TargetShape { Info = info, Com = s });
                sheet.Shapes.Add(info);
                return true;
            }

            return false;
        }

        /// <summary>名前が "_" で始まる、または非表示の図形は対象外。</summary>
        private static bool IsExcluded(XL.Shape s)
        {
            var name = s.Name ?? "";
            if (name.StartsWith("_", StringComparison.Ordinal)) return true;
            return s.Visible == Office.MsoTriState.msoFalse;
        }

        /// <summary>左上のセルが非表示の行・列にあるなら対象外。</summary>
        private static bool IsOnHiddenCells(XL.Shape s)
        {
            using (var com = new ComScope())
            {
                try
                {
                    var cell = com.Add(s.TopLeftCell);
                    var row = com.Add(cell.EntireRow);
                    var col = com.Add(cell.EntireColumn);
                    return true.Equals(row.Hidden) || true.Equals(col.Hidden);
                }
                catch (COMException)
                {
                    return false;
                }
            }
        }

        private static RectD Bounds(XL.Shape s) => new RectD(s.Left, s.Top, s.Width, s.Height);

        private static string SafeName(XL.Shape s)
        {
            try { return s.Name; } catch { return "?"; }
        }

        private static void ReadText(XL.Shape s, ShapeInfo info)
        {
            using (var com = new ComScope())
            {
                var frame = com.Add(s.TextFrame2);
                var range = com.Add(frame.TextRange);
                info.Text = range.Text ?? "";

                // 書式が混在していると Font.Size は不定値を返すので、先頭の文字のサイズを使う
                double size = 0;
                try
                {
                    var first = com.Add(info.Text.Length > 0 ? range.Characters[1, 1] : range);
                    var font = com.Add(first.Font);
                    size = font.Size;
                }
                catch (COMException)
                {
                }
                info.FontSize = size > 0 && size < 1000 ? size : 11;

                try
                {
                    var para = com.Add(info.Text.Length > 0 ? range.Paragraphs[1, 1] : range);
                    var format = com.Add(para.ParagraphFormat);
                    switch (format.Alignment)
                    {
                        case Office.MsoParagraphAlignment.msoAlignCenter:
                        case Office.MsoParagraphAlignment.msoAlignDistribute:
                            info.Alignment = HAlign.Center;
                            break;
                        case Office.MsoParagraphAlignment.msoAlignRight:
                            info.Alignment = HAlign.Right;
                            break;
                        default:
                            info.Alignment = HAlign.Left;
                            break;
                    }
                }
                catch (COMException)
                {
                    info.Alignment = HAlign.Left;
                }
            }
        }

        private static void ReadControlState(XL.Worksheet ws, XL.Shape s, ShapeInfo info)
        {
            using (var com = new ComScope())
            {
                var cf = com.Add(s.ControlFormat);
                double value = 0;
                try { value = Convert.ToDouble(cf.Value); } catch { }
                info.IsChecked = Math.Abs(value - XlOn) < 0.001;

                string link = null;
                try { link = cf.LinkedCell; } catch { }
                if (string.IsNullOrWhiteSpace(link)) return;

                var linked = ResolveRange(ws, link, com);
                if (linked == null)
                {
                    info.LinkedCell = link.Replace("$", "").ToUpperInvariant();
                    return;
                }

                var linkedSheet = com.Add((XL.Worksheet)linked.Worksheet);
                string address = linked.Address[false, false, XL.XlReferenceStyle.xlA1];
                info.LinkedCell = linkedSheet.Name + "!" + address;

                // チェックボックスはリンクセルの値（TRUE/FALSE）を優先
                if (info.Kind == TargetKind.CheckBox)
                {
                    var v = linked.Value2;
                    if (v is bool b) info.IsChecked = b;
                }
            }
        }

        /// <summary>LinkedCell の文字列（A1 形式または R1C1 形式、シート名付きもあり）から Range を得る。</summary>
        internal static XL.Range ResolveRange(XL.Worksheet ws, string address, ComScope com)
        {
            XL.Application app = null;
            try
            {
                app = (XL.Application)ws.Application;
                var r = TryRange(ws, app, address, com);
                if (r != null) return r;

                // 参照形式が R1C1 の場合は A1 形式に変換して再試行
                string a1 = null;
                try
                {
                    a1 = app.ConvertFormula(address, XL.XlReferenceStyle.xlR1C1, XL.XlReferenceStyle.xlA1, Type.Missing, Type.Missing) as string;
                }
                catch (COMException)
                {
                }
                return string.IsNullOrEmpty(a1) ? null : TryRange(ws, app, a1, com);
            }
            finally
            {
                ComScope.Release(app);
            }
        }

        private static XL.Range TryRange(XL.Worksheet ws, XL.Application app, string address, ComScope com)
        {
            try
            {
                // シート名付き（別シート）の参照は Worksheet.Range では解決できないので Application.Range を使う
                return address.Contains("!") ? com.Add(app.Range[address]) : com.Add(ws.Range[address]);
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
