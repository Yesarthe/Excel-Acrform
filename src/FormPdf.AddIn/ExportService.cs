using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FormPdf.AddIn.Excel;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Layout;
using FormPdf.Core.Model;
using FormPdf.Core.Pdf;
using Office = Microsoft.Office.Core;
using XL = Microsoft.Office.Interop.Excel;

namespace FormPdf.AddIn
{
    /// <summary>「フォームPDF出力」ボタンの処理本体。</summary>
    internal sealed class ExportService
    {
        private readonly XL.Application _app;
        private readonly IPdfFormWriter _writer;

        public ExportService(XL.Application app, IPdfFormWriter writer)
        {
            _app = app;
            _writer = writer;
        }

        private sealed class SheetJob
        {
            public object Com;              // Worksheet または Chart 等
            public XL.Worksheet Worksheet;  // ワークシートでなければ null
            public SheetInfo Info;
            public List<TargetShape> Targets = new List<TargetShape>();
        }

        public void Run(FormPdfSettings settings)
        {
            XL.Workbook wb = null;
            var jobs = new List<SheetJob>();
            var warnings = new WarningList();
            var owner = new ExcelWindow(_app);

            try
            {
                wb = _app.ActiveWorkbook;
                if (wb == null)
                {
                    Dialogs.Info(owner, "ブックが開かれていません。");
                    return;
                }

                jobs = CollectSheets(wb, settings.Scope);
                if (jobs.Count == 0)
                {
                    Dialogs.Info(owner, "出力するシートがありません。");
                    return;
                }

                string output = AskOutputPath(owner, wb);
                if (output == null) return;

                RunCore(wb, jobs, output, settings, warnings, owner);
            }
            finally
            {
                foreach (var job in jobs)
                {
                    foreach (var t in job.Targets) ComScope.Release(t.Com);
                    ComScope.Release(job.Com);
                }
                ComScope.Release(wb);
            }
        }

        private void RunCore(XL.Workbook wb, List<SheetJob> jobs, string output, FormPdfSettings settings,
            WarningList warnings, IWin32Window owner)
        {
            var sw = Stopwatch.StartNew();
            string tempPdf = Path.Combine(Path.GetTempPath(), "FormPdf_" + Guid.NewGuid().ToString("N") + ".pdf");

            bool screenUpdating = _app.ScreenUpdating;
            bool enableEvents = _app.EnableEvents;
            bool saved = wb.Saved;
            int fieldCount;
            string markerPdf = null;
            var markerShapes = new List<ShapeInfo>();

            try
            {
                _app.ScreenUpdating = false;
                _app.EnableEvents = false;

                // 1. 図形の走査
                foreach (var job in jobs.Where(j => j.Worksheet != null))
                {
                    Status($"シート「{job.Info.Name}」の図形を読み取り中…");
                    ShapeScanner.Scan(job.Worksheet, job.Info, job.Targets, warnings);
                }

                int targetCount = jobs.Sum(j => j.Targets.Count);
                if (targetCount == 0)
                {
                    _app.ScreenUpdating = screenUpdating;
                    var answer = Dialogs.YesNo(owner,
                        "変換対象（テキストボックス・チェックボックス・オプションボタン）が見つかりません。\n\n" +
                        "通常の PDF として出力しますか？\n［はい］通常の PDF を出力　［いいえ］中止");
                    if (answer != DialogResult.Yes) return;
                    _app.ScreenUpdating = false;
                }

                // 2. 図形の状態を一時変更 → ページ設定読み取り → Excel で PDF 出力 → 必ず元に戻す
                var guards = new List<ShapeStateGuard>();
                try
                {
                    foreach (var job in jobs.Where(j => j.Worksheet != null))
                    {
                        var guard = new ShapeStateGuard(job.Info.Name, warnings);
                        guards.Add(guard);
                        foreach (var t in job.Targets) guard.Apply(job.Worksheet, t);
                    }

                    foreach (var job in jobs)
                    {
                        Status($"シート「{job.Info.Name}」のページ設定を読み取り中…");
                        if (job.Worksheet != null && job.Targets.Count > 0)
                            job.Info.Print = PrintSetupReader.Read(_app, wb, job.Worksheet, warnings);
                        job.Info.ExcelPageCount = PrintSetupReader.GetPageCount(_app, wb, job.Com);
                    }

                    if (settings.AutoCalibrate && targetCount > 0)
                    {
                        Status("図形の位置を測定中…");
                        markerPdf = tempPdf + ".markers.pdf";
                        try
                        {
                            ExportMarkers(wb, jobs, settings.Scope, markerPdf, markerShapes, warnings);
                        }
                        catch (Exception ex)
                        {
                            Logger.Write("自動位置合わせ用の出力に失敗: " + ex);
                            warnings.Add("自動位置合わせができなかったため、計算で配置しました: " + ex.Message);
                            TryDelete(markerPdf);
                            markerPdf = null;
                        }
                    }

                    Status("PDF を出力中…");
                    ExportWithExcel(wb, jobs, settings.Scope, tempPdf);
                }
                finally
                {
                    for (int i = guards.Count - 1; i >= 0; i--) guards[i].Dispose();
                }
            }
            finally
            {
                try { wb.Saved = saved; } catch (COMException) { }
                _app.EnableEvents = enableEvents;
                _app.ScreenUpdating = screenUpdating;
            }

            try
            {
                // 3. フィールドを追加
                Status("フォームフィールドを作成中…");
                var spec = new FormDocumentSpec();
                if (jobs.Sum(j => j.Targets.Count) > 0)
                {
                    var pageSizes = _writer.ReadPageSizes(tempPdf);
                    if (pageSizes.Any(p => p.Rotation != 0))
                        warnings.Add("回転したページがあります。フィールドの位置がずれる可能性があります。");
                    if (markerPdf != null)
                        ApplyMeasurements(markerPdf, markerShapes, pageSizes, warnings);
                    spec = new FormSpecBuilder(settings, warnings).Build(jobs.Select(j => j.Info).ToList(), pageSizes);
                }

                fieldCount = WriteOutput(tempPdf, output, spec, settings, warnings, owner);
                if (fieldCount < 0) return;
            }
            finally
            {
                TryDelete(tempPdf);
                if (markerPdf != null) TryDelete(markerPdf);
                _app.StatusBar = false;
            }

            sw.Stop();
            Logger.Write(
                $"出力: ブック={wb.Name} シート={string.Join(",", jobs.Select(j => j.Info.Name))} " +
                $"フィールド={fieldCount} 警告={warnings.Count} 時間={sw.Elapsed.TotalSeconds:0.0}s " +
                $"フォント={(_writer as ITextPdfFormWriter)?.LastFontDescription} 出力先={output}",
                warnings.Items);

            if (warnings.Count > 0)
                Dialogs.Warnings(owner, $"PDF を出力しました（フィールド {fieldCount} 個）。\n次の警告があります:", warnings.Items);

            if (settings.OpenAfterExport)
            {
                try
                {
                    Process.Start(new ProcessStartInfo(output) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    Logger.Write("PDF を開けませんでした: " + ex.Message);
                }
            }
            else if (warnings.Count == 0)
            {
                Dialogs.Info(owner, $"PDF を出力しました（フィールド {fieldCount} 個）。\n{output}");
            }
        }

        /// <summary>
        /// 自動位置合わせ: 各図形と同じ位置・大きさの四角を図形ごとに固有の色で置き、一度 PDF 出力する。
        /// 四角は出力後に必ず削除する。markerShapes[i] がマーカー番号 i の図形。
        /// </summary>
        private void ExportMarkers(XL.Workbook wb, List<SheetJob> jobs, ExportScope scope, string path,
            List<ShapeInfo> markerShapes, WarningList warnings)
        {
            var markers = new List<XL.Shape>();
            try
            {
                foreach (var job in jobs.Where(j => j.Worksheet != null && j.Targets.Count > 0))
                {
                    XL.Shapes shapes = job.Worksheet.Shapes;
                    try
                    {
                        foreach (var t in job.Targets)
                        {
                            if (markerShapes.Count >= MarkerCalibration.Capacity)
                            {
                                warnings.Add($"図形が {MarkerCalibration.Capacity} 個を超えるため、超えた分は計算で配置します。");
                                break;
                            }
                            var b = t.Info.Bounds;
                            var m = shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                                (float)b.Left, (float)b.Top, (float)Math.Max(1, b.Width), (float)Math.Max(1, b.Height));
                            markers.Add(m);
                            using (var com = new ComScope())
                            {
                                m.Name = "_FormPdfMarker" + markerShapes.Count;
                                var line = com.Add(m.Line);
                                line.Visible = Office.MsoTriState.msoFalse;
                                var shadow = com.Add(m.Shadow);
                                shadow.Visible = Office.MsoTriState.msoFalse;
                                var fill = com.Add(m.Fill);
                                fill.Visible = Office.MsoTriState.msoTrue;
                                fill.Solid();
                                fill.Transparency = 0;
                                var fore = com.Add(fill.ForeColor);
                                fore.RGB = MarkerCalibration.ExcelRgbFor(markerShapes.Count);
                            }
                            markerShapes.Add(t.Info);
                        }
                    }
                    finally
                    {
                        ComScope.Release(shapes);
                    }
                }

                ExportWithExcel(wb, jobs, scope, path);
            }
            finally
            {
                foreach (var m in markers)
                {
                    try
                    {
                        m.Delete();
                    }
                    catch (Exception ex)
                    {
                        Logger.Write("位置合わせ用の図形を削除できませんでした: " + ex.Message);
                        warnings.Add("位置合わせ用の図形（名前が _FormPdfMarker で始まる）を削除できませんでした。手動で削除してください。");
                    }
                    ComScope.Release(m);
                }
            }
        }

        private void ApplyMeasurements(string markerPdf, List<ShapeInfo> markerShapes, IList<PageSizeD> pageSizes, WarningList warnings)
        {
            try
            {
                var markerPages = _writer.ReadPageSizes(markerPdf);
                if (markerPages.Count != pageSizes.Count)
                {
                    warnings.Add($"位置測定用の PDF とページ数が異なる（{markerPages.Count} / {pageSizes.Count}）ため、計算で配置しました。");
                    return;
                }
                int n = MarkerCalibration.Apply(_writer.ReadFilledAreas(markerPdf), markerShapes, pageSizes);
                Logger.Write($"自動位置合わせ: {n} / {markerShapes.Count} 個を測定");
                foreach (var s in markerShapes.Where(x => x.MeasuredRect == null))
                    warnings.Add($"{s.Name}: 位置を測定できなかったため計算で配置しました。");
            }
            catch (Exception ex)
            {
                Logger.Write("位置測定に失敗: " + ex);
                warnings.Add("位置を測定できなかったため、計算で配置しました: " + ex.Message);
            }
        }

        /// <returns>フィールド数。書き込めなかったら -1。</returns>
        private int WriteOutput(string tempPdf, string output, FormDocumentSpec spec, FormPdfSettings settings,
            WarningList warnings, IWin32Window owner)
        {
            // 一時ファイルに書いてから置き換える（出力先を開いていても元のファイルを壊さない）
            string tempOut = tempPdf + ".out.pdf";
            try
            {
                int count = 0;
                if (spec.Count > 0)
                    count = _writer.Write(tempPdf, tempOut, spec, settings, warnings);
                else
                    File.Copy(tempPdf, tempOut, true);

                if (File.Exists(output)) File.Delete(output);
                File.Move(tempOut, output);
                return count;
            }
            catch (IOException ex)
            {
                Logger.Write("出力先に書き込めません: " + ex);
                Dialogs.Error(owner, $"PDF を保存できませんでした。\n出力先の PDF を開いている場合は閉じてから再実行してください。\n\n{ex.Message}");
                return -1;
            }
            catch (UnauthorizedAccessException ex)
            {
                Logger.Write("出力先に書き込めません: " + ex);
                Dialogs.Error(owner, $"PDF を保存できませんでした（アクセス権がありません）。\n\n{ex.Message}");
                return -1;
            }
            finally
            {
                TryDelete(tempOut);
            }
        }

        private List<SheetJob> CollectSheets(XL.Workbook wb, ExportScope scope)
        {
            var list = new List<object>();
            switch (scope)
            {
                case ExportScope.ActiveSheet:
                    list.Add(wb.ActiveSheet);
                    break;
                case ExportScope.SelectedSheets:
                {
                    var window = _app.ActiveWindow;
                    var selected = window.SelectedSheets;
                    foreach (var s in selected) list.Add(s);
                    ComScope.Release(selected);
                    ComScope.Release(window);
                    break;
                }
                case ExportScope.Workbook:
                {
                    var sheets = wb.Sheets;
                    foreach (var s in sheets)
                    {
                        if (IsVisible(s)) list.Add(s);
                        else ComScope.Release(s);
                    }
                    ComScope.Release(sheets);
                    break;
                }
            }

            // PDF にはシート見出しの順に出力される
            return list
                .Select(o => new SheetJob
                {
                    Com = o,
                    Worksheet = o as XL.Worksheet,
                    Info = new SheetInfo { Name = SheetName(o), IsWorksheet = o is XL.Worksheet },
                })
                .OrderBy(j => SheetIndex(j.Com))
                .ToList();
        }

        private static bool IsVisible(object sheet)
        {
            try
            {
                dynamic d = sheet;
                return (XL.XlSheetVisibility)d.Visible == XL.XlSheetVisibility.xlSheetVisible;
            }
            catch
            {
                return true;
            }
        }

        private static string SheetName(object sheet)
        {
            try { dynamic d = sheet; return (string)d.Name; } catch { return "Sheet"; }
        }

        private static int SheetIndex(object sheet)
        {
            try { dynamic d = sheet; return (int)d.Index; } catch { return 0; }
        }

        /// <summary>Excel 標準の「PDF として保存」で出力する。</summary>
        private void ExportWithExcel(XL.Workbook wb, List<SheetJob> jobs, ExportScope scope, string path)
        {
            if (scope == ExportScope.Workbook)
            {
                wb.ExportAsFixedFormat(XL.XlFixedFormatType.xlTypePDF, path,
                    XL.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, false, Type.Missing);
                return;
            }

            // Worksheet.ExportAsFixedFormat はグループ選択中のシートをすべて出力するため、
            // 出力対象だけを選択した状態にしてから呼び、選択状態を元に戻す。
            var window = _app.ActiveWindow;
            var originalActive = wb.ActiveSheet;
            var originalSelection = new List<object>();
            var selected = window.SelectedSheets;
            foreach (var s in selected) originalSelection.Add(s);
            ComScope.Release(selected);

            try
            {
                bool first = true;
                foreach (var job in jobs)
                {
                    dynamic sheet = job.Com;
                    sheet.Select(first);
                    first = false;
                }

                // アクティブシートは常に出力対象に含まれる（アクティブシート／選択中のシートのどちらでも）
                dynamic active = originalActive;
                active.ExportAsFixedFormat(XL.XlFixedFormatType.xlTypePDF, path,
                    XL.XlFixedFormatQuality.xlQualityStandard, true, false, Type.Missing, Type.Missing, false, Type.Missing);
            }
            finally
            {
                try
                {
                    dynamic act = originalActive;
                    act.Select(true);
                    foreach (var s in originalSelection)
                    {
                        if (SheetIndex(s) == SheetIndex(originalActive)) continue;
                        dynamic d = s;
                        d.Select(false);
                    }
                    act.Activate();
                }
                catch (Exception ex)
                {
                    Logger.Write("シートの選択状態を戻せませんでした: " + ex.Message);
                }
                foreach (var s in originalSelection) ComScope.Release(s);
                ComScope.Release(originalActive);
                ComScope.Release(window);
            }
        }

        private string AskOutputPath(IWin32Window owner, XL.Workbook wb)
        {
            string dir = null;
            try { dir = wb.Path; } catch { }
            if (string.IsNullOrEmpty(dir) || dir.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            using (var dlg = new SaveFileDialog
            {
                Title = "フォームPDFの保存先",
                Filter = "PDF ファイル (*.pdf)|*.pdf",
                DefaultExt = "pdf",
                AddExtension = true,
                OverwritePrompt = true,
                InitialDirectory = dir,
                FileName = Path.GetFileNameWithoutExtension(wb.Name) + ".pdf",
            })
            {
                return dlg.ShowDialog(owner) == DialogResult.OK ? dlg.FileName : null;
            }
        }

        private void Status(string message)
        {
            try { _app.StatusBar = "フォームPDF: " + message; } catch (COMException) { }
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }
    }
}
