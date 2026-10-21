using System;
using System.Collections.Generic;
using System.Linq;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Layout;
using FormPdf.Core.Model;

namespace FormPdf.Core.Fields
{
    /// <summary>
    /// シート情報 + 出力 PDF のページサイズから、作成するフィールド一覧を組み立てる。
    /// COM にも PDF ライブラリにも依存しない（単体テスト可能）。
    /// </summary>
    public sealed class FormSpecBuilder
    {
        private readonly FormPdfSettings _settings;
        private readonly WarningList _warnings;

        private Dictionary<int, List<PdfBox>> _drawingsByPage;

        public FormSpecBuilder(FormPdfSettings settings, WarningList warnings)
        {
            _settings = settings;
            _warnings = warnings;
        }

        /// <summary>
        /// 出力 PDF の描画（IPdfFormWriter.ReadDrawnAreas の結果）。
        /// 設定されていれば、チェックボックスの四角・オプションボタンの丸を印字から探して重ねる。
        /// </summary>
        public IList<FilledArea> PageDrawings
        {
            set => _drawingsByPage = value?.GroupBy(a => a.PageIndex).ToDictionary(g => g.Key, g => g.Select(a => a.Bounds).ToList());
        }

        /// <summary>チェックボックス・オプションボタンの数と、そのうち印字から位置を検出できた数（ログ用）。</summary>
        public int ButtonCount { get; private set; }
        public int DetectedButtonCount { get; private set; }
        /// <summary>ボタンごとの配置の記録（ページ、コントロール全体、ウィジェット、検出の有無）。ログ用。</summary>
        public List<string> ButtonLog { get; } = new List<string>();

        /// <param name="sheets">PDF に出力された順のシート。</param>
        /// <param name="pageSizes">出力 PDF の全ページのサイズ。</param>
        public FormDocumentSpec Build(IList<SheetInfo> sheets, IList<PageSizeD> pageSizes)
        {
            var spec = new FormDocumentSpec();
            // 複数シートを出力するときは「シート名.図形名」の階層名にする
            bool multiSheet = sheets.Count(s => s.IsWorksheet) > 1;

            var topNames = new NameAllocator();
            int pageStart = 0;

            foreach (var sheet in sheets)
            {
                SheetLayout layout = null;
                int pageCount;

                if (sheet.IsWorksheet && sheet.Print.Areas.Count > 0)
                {
                    var firstSize = pageStart < pageSizes.Count ? pageSizes[pageStart] : new PageSizeD(595.3, 841.9);
                    layout = PrintLayoutCalculator.Compute(sheet.Print, firstSize, _settings.ScaleCorrectionPercent);
                    pageCount = layout.Pages.Count;
                    if (sheet.ExcelPageCount.HasValue && sheet.ExcelPageCount.Value != pageCount)
                    {
                        // 全図形を実測できていれば計算上のページ数は使わないので警告しない
                        if (!AllMeasured(sheet.Shapes))
                            _warnings.Add(sheet.Name, null,
                            $"計算したページ数（{pageCount}）と Excel のページ数（{sheet.ExcelPageCount}）が一致しません。位置がずれる可能性があります。");
                        pageCount = sheet.ExcelPageCount.Value;
                    }
                    // 実際のページサイズを各ページに反映（シート内で用紙が変わることは無いが念のため）
                    foreach (var p in layout.Pages)
                    {
                        int pdfIndex = pageStart + p.IndexInSheet;
                        if (pdfIndex < pageSizes.Count)
                            p.PageSize = pageSizes[pdfIndex];
                    }
                }
                else
                {
                    pageCount = sheet.ExcelPageCount ?? (sheet.IsWorksheet ? 0 : 1);
                }

                if (sheet.IsWorksheet && sheet.Shapes.Count > 0)
                {
                    string parent = null;
                    if (multiSheet)
                        parent = topNames.Allocate(NameAllocator.Sanitize(sheet.Name));
                    var names = multiSheet ? new NameAllocator() : topNames;
                    if (layout == null && sheet.Shapes.Any(x => x.MeasuredRect == null))
                        _warnings.Add(sheet.Name, null, "印刷範囲が取得できないため、位置を測定できなかった図形はスキップします。");
                    BuildSheet(spec, sheet, layout, pageStart, pageSizes.Count, parent, names);
                }

                pageStart += pageCount;
            }

            if (pageStart != pageSizes.Count && pageSizes.Count > 0 && !AllMeasured(sheets.SelectMany(x => x.Shapes)))
            {
                _warnings.Add($"PDF のページ数（{pageSizes.Count}）と計算上のページ数（{pageStart}）が一致しません。位置がずれている可能性があります。");
            }
            ApplyLearnedToUndetected();
            return spec;
        }

        private static bool AllMeasured(IEnumerable<ShapeInfo> shapes) =>
            shapes.All(x => x.MeasuredRect.HasValue && x.MeasuredPageIndex.HasValue);

        private void BuildSheet(FormDocumentSpec spec, SheetInfo sheet, SheetLayout layout,
            int pageStart, int pdfPageCount, string parent, NameAllocator names)
        {
            var mapped = new Dictionary<ShapeInfo, Placement>();
            foreach (var shape in sheet.Shapes)
            {
                Placement p;
                if (shape.MeasuredRect.HasValue && shape.MeasuredPageIndex.HasValue)
                {
                    // 自動位置合わせで実測した位置（ページも実測値）
                    var r = shape.MeasuredRect.Value;
                    p = new Placement
                    {
                        Page = shape.MeasuredPageIndex.Value,
                        Rect = new PdfBox(r.X + _settings.OffsetX, r.Y - _settings.OffsetY, r.Width, r.Height),
                        Scale = shape.Bounds.Width > 1 ? r.Width / shape.Bounds.Width
                              : shape.Bounds.Height > 1 ? r.Height / shape.Bounds.Height : 1,
                    };
                }
                else
                {
                    var m = layout == null ? null : PrintLayoutCalculator.Map(layout, shape.Bounds, _settings.OffsetX, _settings.OffsetY);
                    if (m == null)
                    {
                        _warnings.Add(sheet.Name, shape.Name, "印刷範囲の外にあるためスキップしました。");
                        continue;
                    }
                    p = new Placement { Page = pageStart + m.Value.PageIndexInSheet, Rect = m.Value.PdfRect, Scale = m.Value.Scale };
                }

                if (p.Page < 0 || p.Page >= pdfPageCount)
                {
                    _warnings.Add(sheet.Name, shape.Name, "対応する PDF ページが見つからないためスキップしました。");
                    continue;
                }
                mapped[shape] = p;
            }

            // テキストボックス・チェックボックス（図形の並び順）
            foreach (var shape in sheet.Shapes.Where(s => mapped.ContainsKey(s)))
            {
                var m = mapped[shape];
                switch (shape.Kind)
                {
                    case TargetKind.TextBox:
                        spec.Fields.Add(BuildText(shape, m, sheet, parent, names));
                        break;
                    case TargetKind.CheckBox:
                        var cb = new CheckBoxSpec
                        {
                            ParentName = parent,
                            Name = names.Allocate(NameAllocator.Sanitize(shape.Name)),
                            SourceSheet = sheet.Name,
                            SourceShape = shape.Name,
                            PageIndex = m.Page,
                            Checked = _settings.InheritCheckStates && shape.IsChecked,
                        };
                        cb.Rect = ButtonRect(m, _settings.CheckBoxSize, TargetKind.CheckBox, shape.Name, r => cb.Rect = r);
                        spec.Fields.Add(cb);
                        break;
                }
            }

            // ラジオボタン
            var options = sheet.Shapes.Where(s => s.Kind == TargetKind.OptionButton && mapped.ContainsKey(s)).ToList();
            foreach (var group in GroupOptionButtons(sheet, options))
            {
                var g = new RadioGroupSpec
                {
                    ParentName = parent,
                    Name = names.Allocate(NameAllocator.Sanitize(group.Name)),
                    SourceSheet = sheet.Name,
                    SourceShape = group.Name,
                };
                var exportNames = new NameAllocator();
                foreach (var shape in group.Members)
                {
                    var m = mapped[shape];
                    var b = new RadioButtonSpec
                    {
                        ExportValue = exportNames.Allocate(shape.Name),
                        PageIndex = m.Page,
                        SourceShape = shape.Name,
                    };
                    b.Rect = ButtonRect(m, _settings.OptionButtonSize, TargetKind.OptionButton, shape.Name, r => b.Rect = r);
                    g.Buttons.Add(b);
                    if (_settings.InheritCheckStates && shape.IsChecked && g.SelectedValue == null)
                        g.SelectedValue = b.ExportValue;
                }
                spec.Fields.Add(g);
            }
        }

        private TextFieldSpec BuildText(ShapeInfo shape, Placement m, SheetInfo sheet, string parent, NameAllocator names)
        {
            double fontSize = shape.FontSize > 0 ? shape.FontSize : 11;
            return new TextFieldSpec
            {
                ParentName = parent,
                Name = names.Allocate(NameAllocator.Sanitize(shape.Name)),
                SourceSheet = sheet.Name,
                SourceShape = shape.Name,
                PageIndex = m.Page,
                Rect = m.Rect,
                Value = _settings.InheritTextValues ? NormalizeNewLines(shape.Text) : "",
                FontSize = _settings.ScaleFontWithPage ? fontSize * m.Scale : fontSize,
                Alignment = shape.Alignment,
                Multiline = shape.Bounds.Height >= fontSize * 2,
            };
        }

        /// <summary>印字から検出できたボタンの、コントロールに対する四角・丸の位置（印刷倍率で割った値）。</summary>
        private readonly Dictionary<TargetKind, List<(double dx, double dy, double size)>> _learned =
            new Dictionary<TargetKind, List<(double, double, double)>>();

        /// <summary>検出できなかったボタン。Build の最後に、検出できたボタンの位置関係を当てはめる。</summary>
        private readonly List<(TargetKind kind, Placement m, string name, Action<PdfBox> set)> _undetected =
            new List<(TargetKind, Placement, string, Action<PdfBox>)>();

        /// <summary>
        /// ボタン部分（四角・丸）の位置。Excel の印字から見つかればそれに合わせる。
        /// 見つからなければいったん図形の左端・縦中央から正方形を切り出し、
        /// Build の最後に、同じ PDF で検出できたボタンの位置関係があればそれで置き直す。
        /// </summary>
        private PdfBox ButtonRect(Placement m, double sizeSetting, TargetKind kind, string name, Action<PdfBox> setLater)
        {
            double s = m.Scale;
            ButtonCount++;
            if (_settings.DetectButtonFrames && _drawingsByPage != null && _drawingsByPage.TryGetValue(m.Page, out var drawings))
            {
                double expected = (sizeSetting > 0 ? sizeSetting : 10) * s;
                var found = ButtonFrameDetector.Find(drawings, m.Rect, expected);
                if (found.HasValue)
                {
                    var f = found.Value;
                    DetectedButtonCount++;
                    ButtonLog.Add($"{name} p{m.Page + 1} 全体={m.Rect} → 検出={f}");
                    if (s > 0)
                    {
                        if (!_learned.TryGetValue(kind, out var list)) _learned[kind] = list = new List<(double, double, double)>();
                        list.Add(((f.X - m.Rect.X) / s,
                                  (f.Y + f.Height / 2 - (m.Rect.Y + m.Rect.Height / 2)) / s,
                                  Math.Max(f.Width, f.Height) / s));
                    }
                    return f;
                }
            }

            double full = m.Rect.Height;
            double size = sizeSetting > 0 ? Math.Min(sizeSetting * s, full) : full;
            size = Math.Min(size, m.Rect.Width);
            double x = m.Rect.X + _settings.ButtonLeftInset * s;
            double y = m.Rect.Y + (full - size) / 2;
            var box = new PdfBox(x, y, size, size);
            _undetected.Add((kind, m, name, setLater));
            return box;
        }

        /// <summary>
        /// 検出できなかったボタンを、同じ種類（無ければもう一方の種類）の検出結果の中央値で置き直す。
        /// Excel の四角・丸は大きさも左端からの位置もコントロールによらず一定なので、1 個でも検出できれば合わせられる。
        /// </summary>
        private void ApplyLearnedToUndetected()
        {
            foreach (var (kind, m, name, set) in _undetected)
            {
                if (!_learned.TryGetValue(kind, out var list) || list.Count == 0)
                    list = _learned.Values.FirstOrDefault(l => l.Count > 0);
                if (list == null)
                {
                    ButtonLog.Add($"{name} p{m.Page + 1} 全体={m.Rect} → 既定（検出例なし）");
                    continue;
                }
                double s = m.Scale;
                double dx = Median(list.Select(t => t.dx)) * s;
                double dy = Median(list.Select(t => t.dy)) * s;
                double size = Median(list.Select(t => t.size)) * s;
                double cy = m.Rect.Y + m.Rect.Height / 2 + dy;
                var box = new PdfBox(m.Rect.X + dx, cy - size / 2, size, size);
                set(box);
                ButtonLog.Add($"{name} p{m.Page + 1} 全体={m.Rect} → 推定={box}（検出例 {list.Count} 個から）");
            }
            _undetected.Clear();
        }

        private static double Median(IEnumerable<double> values)
        {
            var v = values.OrderBy(x => x).ToList();
            int n = v.Count;
            return n % 2 == 1 ? v[n / 2] : (v[n / 2 - 1] + v[n / 2]) / 2;
        }

        /// <summary>図形 1 個の配置先（PDF 全体でのページ番号、PDF 座標、印刷倍率）。</summary>
        private struct Placement
        {
            public int Page;
            public PdfBox Rect;
            public double Scale;
        }

        private static string NormalizeNewLines(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            // Excel のテキストボックスは段落区切りに CR（\r）、改行に VT（\v）を使う
            return text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n');
        }

        public sealed class OptionGroup
        {
            public string Name { get; set; }
            public List<ShapeInfo> Members { get; } = new List<ShapeInfo>();
        }

        /// <summary>
        /// オプションボタンのグループ分け。
        /// 1) リンクセルが同じもの（最優先） 2) 同じグループボックス内 3) それ以外は「シート名_Default」。
        /// </summary>
        public static List<OptionGroup> GroupOptionButtons(SheetInfo sheet, IList<ShapeInfo> options)
        {
            var groups = new List<OptionGroup>();
            var byKey = new Dictionary<string, OptionGroup>(StringComparer.OrdinalIgnoreCase);

            foreach (var shape in options)
            {
                var box = FindGroupBox(sheet.GroupBoxes, shape.Bounds);
                string key;
                string name;
                if (!string.IsNullOrEmpty(shape.LinkedCell))
                {
                    key = "link:" + shape.LinkedCell;
                    name = box?.Name ?? sheet.Name + "_" + shape.LinkedCell.Replace("!", "_");
                }
                else if (box != null)
                {
                    key = "box:" + box.Name + "@" + box.Bounds;
                    name = box.Name;
                }
                else
                {
                    key = "default";
                    name = sheet.Name + "_Default";
                }

                // リンクセルを持たないボタンでも、同じグループボックス内のボタンがリンクセルを持っていればそちらに合流する
                if (!byKey.TryGetValue(key, out var group) && key.StartsWith("box:"))
                {
                    group = groups.FirstOrDefault(g => g.Members.Any(m =>
                        !string.IsNullOrEmpty(m.LinkedCell) && ReferenceEquals(FindGroupBox(sheet.GroupBoxes, m.Bounds), box)));
                    if (group != null) byKey[key] = group;
                }

                if (group == null)
                {
                    group = new OptionGroup { Name = name };
                    byKey[key] = group;
                    groups.Add(group);
                }
                group.Members.Add(shape);
            }
            return groups;
        }

        /// <summary>ボタンの中心を含むグループボックスのうち、最も小さいもの。</summary>
        public static GroupBoxInfo FindGroupBox(IList<GroupBoxInfo> boxes, RectD bounds)
        {
            GroupBoxInfo best = null;
            foreach (var b in boxes)
            {
                if (b.Bounds.Contains(bounds.CenterX, bounds.CenterY) && (best == null || b.Bounds.Area < best.Bounds.Area))
                    best = b;
            }
            return best;
        }
    }

    /// <summary>同じ親の下で重複しない名前を払い出す（名前, 名前_2, 名前_3 ...）。</summary>
    public sealed class NameAllocator
    {
        private readonly HashSet<string> _used = new HashSet<string>(StringComparer.Ordinal);

        public string Allocate(string baseName)
        {
            if (string.IsNullOrEmpty(baseName)) baseName = "Field";
            if (_used.Add(baseName)) return baseName;
            for (int i = 2; ; i++)
            {
                var candidate = baseName + "_" + i;
                if (_used.Add(candidate)) return candidate;
            }
        }

        /// <summary>PDF の部分フィールド名にピリオドは使えないので置き換える。</summary>
        public static string Sanitize(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Field";
            return name.Replace('.', '_').Trim();
        }
    }
}
