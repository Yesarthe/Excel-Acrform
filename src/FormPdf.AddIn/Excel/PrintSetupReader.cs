using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;
using XL = Microsoft.Office.Interop.Excel;

namespace FormPdf.AddIn.Excel
{
    /// <summary>ページ設定・印刷範囲・改ページ位置・ページ数を読み取る。</summary>
    internal static class PrintSetupReader
    {
        private const int MaxRows = 1048576;
        private const int MaxColumns = 16384;

        public static PrintSetupInfo Read(XL.Application app, XL.Workbook wb, XL.Worksheet ws, WarningList warnings)
        {
            var info = new PrintSetupInfo();
            using (var com = new ComScope())
            {
                var ps = com.Add(ws.PageSetup);

                info.LeftMargin = ps.LeftMargin;
                info.RightMargin = ps.RightMargin;
                info.TopMargin = ps.TopMargin;
                info.BottomMargin = ps.BottomMargin;
                info.CenterHorizontally = ps.CenterHorizontally;
                info.CenterVertically = ps.CenterVertically;
                info.Order = ps.Order == XL.XlOrder.xlOverThenDown ? PageOrder.OverThenDown : PageOrder.DownThenOver;

                // Zoom は数値か False（= FitToPages 指定）
                object zoom = ps.Zoom;
                if (zoom is bool)
                {
                    info.Zoom = null;
                    info.FitToPagesWide = ToNullableInt(ps.FitToPagesWide);
                    info.FitToPagesTall = ToNullableInt(ps.FitToPagesTall);
                }
                else
                {
                    info.Zoom = Convert.ToDouble(zoom);
                }

                info.Areas = ReadPrintAreas(ws, ps, com, warnings);
                ReadPrintTitles(ws, info, com);
            }

            ReadBreaks(app, wb, ws, info, warnings);
            return info;
        }

        private static int? ToNullableInt(object v)
        {
            if (v == null || v is bool) return null;
            try
            {
                int n = Convert.ToInt32(v);
                return n > 0 ? n : (int?)null;
            }
            catch
            {
                return null;
            }
        }

        private static RectD RectOf(XL.Range r) => new RectD(r.Left, r.Top, r.Width, r.Height);

        private static List<RectD> ReadPrintAreas(XL.Worksheet ws, XL.PageSetup ps, ComScope com, WarningList warnings)
        {
            var result = new List<RectD>();

            XL.Range range = TryGetNamedRange(ws, "Print_Area", com);
            if (range == null)
            {
                string pa = null;
                try { pa = ps.PrintArea; } catch { }
                if (!string.IsNullOrWhiteSpace(pa))
                    range = ShapeScanner.ResolveRange(ws, pa, com);
            }

            if (range != null)
            {
                var areas = com.Add(range.Areas);
                for (int i = 1; i <= areas.Count; i++)
                {
                    var a = com.Add(areas[i]);
                    result.Add(RectOf(a));
                }
                return result;
            }

            // 印刷範囲が無い場合は使用範囲（＋印刷される図形）を印刷する
            warnings.Add(ws.Name, null, "印刷範囲が設定されていないため使用範囲から推定します。位置がずれる場合は印刷範囲を設定してください。");
            var used = com.Add(ws.UsedRange);
            var rect = RectOf(used);
            double left = rect.Left, top = rect.Top, right = rect.Right, bottom = rect.Bottom;
            XL.Shapes shapes = com.Add(ws.Shapes);
            for (int i = 1; i <= shapes.Count; i++)
            {
                var s = com.Add(shapes.Item(i));
                try
                {
                    if (s.Visible == Microsoft.Office.Core.MsoTriState.msoFalse) continue;
                    var br = com.Add(s.BottomRightCell);
                    var tl = com.Add(s.TopLeftCell);
                    left = Math.Min(left, tl.Left);
                    top = Math.Min(top, tl.Top);
                    right = Math.Max(right, br.Left + br.Width);
                    bottom = Math.Max(bottom, br.Top + br.Height);
                }
                catch (COMException)
                {
                }
            }
            result.Add(new RectD(left, top, right - left, bottom - top));
            return result;
        }

        private static void ReadPrintTitles(XL.Worksheet ws, PrintSetupInfo info, ComScope com)
        {
            var titles = TryGetNamedRange(ws, "Print_Titles", com);
            if (titles == null) return;
            var areas = com.Add(titles.Areas);
            for (int i = 1; i <= areas.Count; i++)
            {
                var a = com.Add(areas[i]);
                var cols = com.Add(a.Columns);
                var rows = com.Add(a.Rows);
                if (cols.Count >= MaxColumns)
                    info.TitleRows = RectOf(a);
                else if (rows.Count >= MaxRows)
                    info.TitleColumns = RectOf(a);
            }
        }

        private static XL.Range TryGetNamedRange(XL.Worksheet ws, string name, ComScope com)
        {
            try
            {
                var names = com.Add(ws.Names);
                var n = com.Add(names.Item(name, Type.Missing, Type.Missing));
                return com.Add(n.RefersToRange);
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>
        /// 改ページ位置（自動改ページを含む）。GET.DOCUMENT(64/65) を優先し、ダメなら HPageBreaks/VPageBreaks。
        /// </summary>
        private static void ReadBreaks(XL.Application app, XL.Workbook wb, XL.Worksheet ws, PrintSetupInfo info, WarningList warnings)
        {
            var rows = GetDocumentNumbers(app, wb, ws, 64);
            var cols = GetDocumentNumbers(app, wb, ws, 65);

            if (rows == null || cols == null)
            {
                bool display = false;
                try
                {
                    // 自動改ページの計算を確実に行わせる
                    display = ws.DisplayPageBreaks;
                    ws.DisplayPageBreaks = true;
                }
                catch (COMException)
                {
                }

                try
                {
                    if (rows == null) rows = ReadHBreaks(ws);
                    if (cols == null) cols = ReadVBreaks(ws);
                }
                finally
                {
                    try { ws.DisplayPageBreaks = display; } catch (COMException) { }
                }
            }

            using (var com = new ComScope())
            {
                var allRows = com.Add(ws.Rows);
                foreach (var r in rows ?? new List<int>())
                {
                    if (r < 1 || r > MaxRows) continue;
                    var row = com.Add((XL.Range)allRows[r]);
                    info.RowBreaks.Add(row.Top);
                }
                var allColumns = com.Add(ws.Columns);
                foreach (var c in cols ?? new List<int>())
                {
                    if (c < 1 || c > MaxColumns) continue;
                    var col = com.Add((XL.Range)allColumns[c]);
                    info.ColumnBreaks.Add(col.Left);
                }
            }
        }

        private static List<int> ReadHBreaks(XL.Worksheet ws)
        {
            var list = new List<int>();
            using (var com = new ComScope())
            {
                var breaks = com.Add(ws.HPageBreaks);
                int n = breaks.Count;
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        var b = com.Add(breaks[i]);
                        var loc = com.Add(b.Location);
                        list.Add(loc.Row);
                    }
                    catch (COMException)
                    {
                        // 範囲外の自動改ページで例外になることがある
                    }
                }
            }
            return list;
        }

        private static List<int> ReadVBreaks(XL.Worksheet ws)
        {
            var list = new List<int>();
            using (var com = new ComScope())
            {
                var breaks = com.Add(ws.VPageBreaks);
                int n = breaks.Count;
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        var b = com.Add(breaks[i]);
                        var loc = com.Add(b.Location);
                        list.Add(loc.Column);
                    }
                    catch (COMException)
                    {
                    }
                }
            }
            return list;
        }

        /// <summary>
        /// Excel 4 マクロ関数 GET.DOCUMENT。64 = 水平改ページ直下の行番号、65 = 垂直改ページ右の列番号、50 = 印刷ページ数。
        /// 取得できなければ null。
        /// </summary>
        internal static List<int> GetDocumentNumbers(XL.Application app, XL.Workbook wb, object sheet, int typeNum)
        {
            try
            {
                string sheetName = sheet is XL.Worksheet w ? w.Name : ((XL.Chart)sheet).Name;
                string reference = "'[" + wb.Name.Replace("'", "''") + "]" + sheetName.Replace("'", "''") + "'";
                object result = app.ExecuteExcel4Macro($"GET.DOCUMENT({typeNum},\"{reference}\")");
                var list = new List<int>();
                Collect(result, list);
                return list;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void Collect(object value, List<int> list)
        {
            switch (value)
            {
                case null:
                    return;
                case double d:
                    list.Add((int)Math.Round(d));
                    return;
                case Array arr:
                    foreach (var v in arr) Collect(v, list);
                    return;
                default:
                    // エラー値（#N/A 等 = Int32）は無視
                    return;
            }
        }

        public static int? GetPageCount(XL.Application app, XL.Workbook wb, object sheet)
        {
            var n = GetDocumentNumbers(app, wb, sheet, 50);
            if (n == null || n.Count == 0) return null;
            return n[0];
        }
    }
}
