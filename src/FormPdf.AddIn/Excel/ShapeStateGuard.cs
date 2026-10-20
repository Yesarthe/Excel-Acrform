using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;
using Office = Microsoft.Office.Core;
using XL = Microsoft.Office.Interop.Excel;

namespace FormPdf.AddIn.Excel
{
    /// <summary>
    /// PDF 出力の間だけ図形の状態を変え、Dispose で必ず元に戻す。
    /// ・テキストボックス: 非表示（Excel の枠や文字とフィールドが二重に見えないように）
    /// ・チェックボックス／オプションボタン: 未チェックにする（四角・丸とキャプションは印字し、✓ や ● は PDF 側で描く）
    /// </summary>
    internal sealed class ShapeStateGuard : IDisposable
    {
        private const int XlOff = -4146;

        private readonly List<(XL.Shape shape, Office.MsoTriState visible, string name)> _hidden =
            new List<(XL.Shape, Office.MsoTriState, string)>();
        private readonly List<(XL.Shape shape, int value, string name)> _values =
            new List<(XL.Shape, int, string)>();
        private readonly List<(XL.Range cell, object formula)> _linkedCells = new List<(XL.Range, object)>();
        private readonly WarningList _warnings;
        private readonly string _sheetName;

        public ShapeStateGuard(string sheetName, WarningList warnings)
        {
            _sheetName = sheetName;
            _warnings = warnings;
        }

        public void Apply(XL.Worksheet ws, TargetShape target)
        {
            var s = target.Com;
            try
            {
                if (target.Info.Kind == TargetKind.TextBox)
                {
                    var v = s.Visible;
                    s.Visible = Office.MsoTriState.msoFalse;
                    _hidden.Add((s, v, target.Info.Name));
                    return;
                }

                // リンクセルの元の内容（数式・空欄を含む）を保存してから値を変える
                using (var com = new ComScope())
                {
                    var cf = com.Add(s.ControlFormat);
                    int value = cf.Value;
                    if (value == XlOff) return;

                    string link = null;
                    try { link = cf.LinkedCell; } catch (COMException) { }
                    if (!string.IsNullOrWhiteSpace(link))
                    {
                        var scope = new ComScope();
                        var cell = ShapeScanner.ResolveRange(ws, link, scope);
                        if (cell != null)
                        {
                            // cell は Dispose まで保持する（scope は解放しない）
                            _linkedCells.Add((cell, cell.Formula));
                        }
                    }

                    cf.Value = XlOff;
                    _values.Add((s, value, target.Info.Name));
                }
            }
            catch (Exception ex)
            {
                _warnings.Add(_sheetName, target.Info.Name,
                    "出力前に図形の状態を変更できませんでした（シート保護の可能性）。PDF 上で二重に表示される場合があります: " + ex.Message);
            }
        }

        public void Dispose()
        {
            // 逆順に戻す
            for (int i = _values.Count - 1; i >= 0; i--)
            {
                var (shape, value, name) = _values[i];
                try
                {
                    using (var com = new ComScope())
                    {
                        var cf = com.Add(shape.ControlFormat);
                        cf.Value = value;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Write($"[{_sheetName}] {name} の値を戻せませんでした: {ex.Message}");
                    _warnings.Add(_sheetName, name, "チェック状態を元に戻せませんでした。確認してください。");
                }
            }

            for (int i = _linkedCells.Count - 1; i >= 0; i--)
            {
                var (cell, formula) = _linkedCells[i];
                try
                {
                    cell.Formula = formula;
                }
                catch (Exception ex)
                {
                    Logger.Write($"[{_sheetName}] リンクセルを戻せませんでした: {ex.Message}");
                }
                finally
                {
                    ComScope.Release(cell);
                }
            }

            for (int i = _hidden.Count - 1; i >= 0; i--)
            {
                var (shape, visible, name) = _hidden[i];
                try
                {
                    shape.Visible = visible;
                }
                catch (Exception ex)
                {
                    Logger.Write($"[{_sheetName}] {name} の表示状態を戻せませんでした: {ex.Message}");
                    _warnings.Add(_sheetName, name, "表示状態を元に戻せませんでした。確認してください。");
                }
            }

            _values.Clear();
            _linkedCells.Clear();
            _hidden.Clear();
        }
    }
}
