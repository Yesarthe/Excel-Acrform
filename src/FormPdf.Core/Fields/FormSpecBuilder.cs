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

        public FormSpecBuilder(FormPdfSettings settings, WarningList warnings)
        {
            _settings = settings;
            _warnings = warnings;
        }

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
                    if (layout == null)
                    {
                        _warnings.Add(sheet.Name, null, "印刷範囲が取得できないため、このシートのフィールドは作成しません。");
                    }
                    else
                    {
                        BuildSheet(spec, sheet, layout, pageStart, pageSizes.Count, parent, names);
                    }
                }

                pageStart += pageCount;
            }

            if (pageStart != pageSizes.Count && pageSizes.Count > 0)
            {
                _warnings.Add($"PDF のページ数（{pageSizes.Count}）と計算上のページ数（{pageStart}）が一致しません。位置がずれている可能性があります。");
            }
            return spec;
        }

        private void BuildSheet(FormDocumentSpec spec, SheetInfo sheet, SheetLayout layout,
            int pageStart, int pdfPageCount, string parent, NameAllocator names)
        {
            var mapped = new Dictionary<ShapeInfo, MappedRect>();
            foreach (var shape in sheet.Shapes)
            {
                var m = PrintLayoutCalculator.Map(layout, shape.Bounds, _settings.OffsetX, _settings.OffsetY);
                if (m == null)
                {
                    _warnings.Add(sheet.Name, shape.Name, "印刷範囲の外にあるためスキップしました。");
                    continue;
                }
                if (pageStart + m.Value.PageIndexInSheet >= pdfPageCount)
                {
                    _warnings.Add(sheet.Name, shape.Name, "対応する PDF ページが見つからないためスキップしました。");
                    continue;
                }
                mapped[shape] = m.Value;
            }

            // テキストボックス・チェックボックス（図形の並び順）
            foreach (var shape in sheet.Shapes.Where(s => mapped.ContainsKey(s)))
            {
                var m = mapped[shape];
                switch (shape.Kind)
                {
                    case TargetKind.TextBox:
                        spec.Fields.Add(BuildText(shape, m, pageStart, sheet, parent, names));
                        break;
                    case TargetKind.CheckBox:
                        spec.Fields.Add(new CheckBoxSpec
                        {
                            ParentName = parent,
                            Name = names.Allocate(NameAllocator.Sanitize(shape.Name)),
                            SourceSheet = sheet.Name,
                            SourceShape = shape.Name,
                            PageIndex = pageStart + m.PageIndexInSheet,
                            Rect = ButtonRect(m, _settings.CheckBoxSize),
                            Checked = _settings.InheritCheckStates && shape.IsChecked,
                        });
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
                        PageIndex = pageStart + m.PageIndexInSheet,
                        Rect = ButtonRect(m, _settings.OptionButtonSize),
                        SourceShape = shape.Name,
                    };
                    g.Buttons.Add(b);
                    if (_settings.InheritCheckStates && shape.IsChecked && g.SelectedValue == null)
                        g.SelectedValue = b.ExportValue;
                }
                spec.Fields.Add(g);
            }
        }

        private TextFieldSpec BuildText(ShapeInfo shape, MappedRect m, int pageStart, SheetInfo sheet, string parent, NameAllocator names)
        {
            double fontSize = shape.FontSize > 0 ? shape.FontSize : 11;
            return new TextFieldSpec
            {
                ParentName = parent,
                Name = names.Allocate(NameAllocator.Sanitize(shape.Name)),
                SourceSheet = sheet.Name,
                SourceShape = shape.Name,
                PageIndex = pageStart + m.PageIndexInSheet,
                Rect = m.PdfRect,
                Value = _settings.InheritTextValues ? NormalizeNewLines(shape.Text) : "",
                FontSize = _settings.ScaleFontWithPage ? fontSize * m.Scale : fontSize,
                Alignment = shape.Alignment,
                Multiline = shape.Bounds.Height >= fontSize * 2,
            };
        }

        /// <summary>図形の左端から、ボタン部分（四角・丸）の正方形を切り出す。</summary>
        private PdfBox ButtonRect(MappedRect m, double sizeSetting)
        {
            double s = m.Scale;
            double full = m.PdfRect.Height;
            double size = sizeSetting > 0 ? Math.Min(sizeSetting * s, full) : full;
            size = Math.Min(size, m.PdfRect.Width);
            double x = m.PdfRect.X + _settings.ButtonLeftInset * s;
            double y = m.PdfRect.Y + (full - size) / 2;
            return new PdfBox(x, y, size, size);
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
