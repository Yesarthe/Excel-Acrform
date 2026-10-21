using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;
using iText.Forms;
using iText.Forms.Fields;
using iText.Forms.Fields.Properties;
using iText.IO.Font;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Extgstate;
using iText.Kernel.Pdf.Xobject;
using iText.Layout.Properties;

namespace FormPdf.Core.Pdf
{
    /// <summary>iText 7 による AcroForm 書き込み。PDF ライブラリ依存はこのクラス（と FontLocator）に閉じ込める。</summary>
    public sealed class ITextPdfFormWriter : IPdfFormWriter
    {
        private readonly FontLocator _fontLocator;

        public ITextPdfFormWriter(FontLocator fontLocator = null)
        {
            _fontLocator = fontLocator ?? new FontLocator();
        }

        /// <summary>直近の Write で使ったフォント（ログ用）。</summary>
        public string LastFontDescription { get; private set; }

        public IList<PageSizeD> ReadPageSizes(string pdfPath)
        {
            var list = new List<PageSizeD>();
            using (var pdf = new PdfDocument(new PdfReader(pdfPath)))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var page = pdf.GetPage(i);
                    var size = page.GetPageSizeWithRotation();
                    list.Add(new PageSizeD(size.GetWidth(), size.GetHeight(), page.GetRotation()));
                }
            }
            return list;
        }

        public IList<FilledArea> ReadFilledAreas(string pdfPath)
        {
            var result = new List<FilledArea>();
            using (var pdf = new PdfDocument(new PdfReader(pdfPath)))
            {
                for (int i = 1; i <= pdf.GetNumberOfPages(); i++)
                {
                    var listener = new FillListener(i - 1, result);
                    new PdfCanvasProcessor(listener).ProcessPageContent(pdf.GetPage(i));
                }
            }
            return result;
        }

        /// <summary>塗りつぶしのパスを拾い、変換行列を適用した外接矩形と塗り色を記録する。</summary>
        private sealed class FillListener : IEventListener
        {
            private readonly int _page;
            private readonly List<FilledArea> _result;

            public FillListener(int page, List<FilledArea> result)
            {
                _page = page;
                _result = result;
            }

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type != EventType.RENDER_PATH || !(data is PathRenderInfo info)) return;
                if ((info.GetOperation() & PathRenderInfo.FILL) == 0) return;
                var color = info.GetFillColor();
                var v = color?.GetColorValue();
                if (v == null || v.Length != 3) return;

                var ctm = info.GetCtm();
                double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
                foreach (var sub in info.GetPath().GetSubpaths())
                    foreach (var seg in sub.GetSegments())
                        foreach (var pt in seg.GetBasePoints())
                        {
                            double x = pt.GetX(), y = pt.GetY();
                            double tx = ctm.Get(Matrix.I11) * x + ctm.Get(Matrix.I21) * y + ctm.Get(Matrix.I31);
                            double ty = ctm.Get(Matrix.I12) * x + ctm.Get(Matrix.I22) * y + ctm.Get(Matrix.I32);
                            minX = Math.Min(minX, tx); maxX = Math.Max(maxX, tx);
                            minY = Math.Min(minY, ty); maxY = Math.Max(maxY, ty);
                        }
                if (minX > maxX) return;

                _result.Add(new FilledArea
                {
                    PageIndex = _page,
                    R = v[0] * 255.0,
                    G = v[1] * 255.0,
                    B = v[2] * 255.0,
                    Bounds = new PdfBox(minX, minY, maxX - minX, maxY - minY),
                });
            }

            public ICollection<EventType> GetSupportedEvents() => new[] { EventType.RENDER_PATH };
        }

        public int Write(string sourcePdf, string outputPdf, FormDocumentSpec spec, FormPdfSettings settings, WarningList warnings)
        {
            int widgets = 0;
            using (var pdf = new PdfDocument(new PdfReader(sourcePdf), new PdfWriter(outputPdf)))
            {
                var form = PdfFormCreator.GetAcroForm(pdf, true);
                // 外観ストリームは自前で生成する（NeedAppearances に頼らない）
                form.SetNeedAppearances(false);

                PdfFont font = spec.Fields.OfType<TextFieldSpec>().Any() || settings.CalibrationMode
                    ? CreateFont(settings, warnings)
                    : null;
                if (font != null)
                    RegisterDefaultResource(form, font);

                var (r, g, b) = settings.GetCheckRgb();
                var checkColor = new DeviceRgb(r, g, b);

                var parents = new Dictionary<string, PdfFormField>(StringComparer.Ordinal);
                var topLevel = new List<PdfFormField>();

                foreach (var field in spec.Fields)
                {
                    try
                    {
                        PdfFormField created;
                        switch (field)
                        {
                            case TextFieldSpec t:
                                if (!CheckPage(pdf, t.PageIndex, t, warnings)) continue;
                                created = CreateTextField(pdf, t, font);
                                widgets++;
                                break;
                            case CheckBoxSpec c:
                                if (!CheckPage(pdf, c.PageIndex, c, warnings)) continue;
                                created = CreateCheckBox(pdf, c, checkColor, form);
                                widgets++;
                                break;
                            case RadioGroupSpec rg:
                                created = CreateRadioGroup(pdf, rg, warnings, out int n);
                                if (created == null) continue;
                                widgets += n;
                                break;
                            default:
                                continue;
                        }

                        if (string.IsNullOrEmpty(field.ParentName))
                        {
                            topLevel.Add(created);
                        }
                        else
                        {
                            if (!parents.TryGetValue(field.ParentName, out var parent))
                            {
                                parent = new NonTerminalFormFieldBuilder(pdf, field.ParentName).CreateNonTerminalFormField();
                                SetUnicodeName(parent, field.ParentName);
                                parents[field.ParentName] = parent;
                                topLevel.Add(parent);
                            }
                            parent.AddKid(created);
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add(field.SourceSheet, field.SourceShape, "フィールドを作成できませんでした: " + ex.Message);
                        Logger.Write($"フィールド作成失敗 {field.FullName}: {ex}");
                    }
                }

                // ウィジェットは各ビルダーで /P（ページ）を設定済み。AddField がそのページに注釈を追加する。
                foreach (var f in topLevel)
                    form.AddField(f, pdf.GetFirstPage());

                if (font != null)
                    RegisterDefaultResource(form, font);

                if (settings.CalibrationMode)
                    DrawCalibration(pdf, spec, font);
            }
            return widgets;
        }

        private static bool CheckPage(PdfDocument pdf, int pageIndex, FieldSpec f, WarningList warnings)
        {
            if (pageIndex >= 0 && pageIndex < pdf.GetNumberOfPages()) return true;
            warnings.Add(f.SourceSheet, f.SourceShape, $"ページ {pageIndex + 1} が PDF に存在しないためスキップしました。");
            return false;
        }

        private static Rectangle ToRect(PdfBox b) =>
            new Rectangle((float)b.X, (float)b.Y, (float)Math.Max(1, b.Width), (float)Math.Max(1, b.Height));

        private static PdfFormField CreateTextField(PdfDocument pdf, TextFieldSpec t, PdfFont font)
        {
            var field = new TextFormFieldBuilder(pdf, t.Name)
                .SetWidgetRectangle(ToRect(t.Rect))
                .SetPage(t.PageIndex + 1)
                .SetFont(font)
                .CreateText();

            SetUnicodeName(field, t.Name);
            field.SetMultiline(t.Multiline);
            field.SetFontSize((float)Math.Max(1, t.FontSize));
            field.SetColor(ColorConstants.BLACK);
            field.SetJustification(t.Alignment == HAlign.Center ? TextAlignment.CENTER
                : t.Alignment == HAlign.Right ? TextAlignment.RIGHT : TextAlignment.LEFT);

            NoBorder(field);
            // 値の設定で外観ストリームが生成される（空でも生成させる）
            field.SetValue(t.Value ?? "");
            return field;
        }

        private PdfFormField CreateCheckBox(PdfDocument pdf, CheckBoxSpec c, DeviceRgb color, PdfAcroForm form)
        {
            var field = new CheckBoxFormFieldBuilder(pdf, c.Name)
                .SetWidgetRectangle(ToRect(c.Rect))
                .SetPage(c.PageIndex + 1)
                .SetCheckType(CheckBoxType.CHECK)
                .CreateCheckBox();

            SetUnicodeName(field, c.Name);
            NoBorder(field);
            field.SetValue(c.Checked ? OnState : "Off");

            // iText はチェックボックスの外観に色を反映しないため、赤い ✓ の外観を自前で作る。
            // /MK /CA '4'（ZapfDingbats の ✓）と /DA の色は、ビューアが外観を作り直す場合の指定。
            RegisterZapfDingbats(form, pdf);
            var rgb = color.GetColorValue();
            var da = $"/{ZapfResourceName} 0 Tf {F(rgb[0])} {F(rgb[1])} {F(rgb[2])} rg";
            field.GetPdfObject().Put(PdfName.DA, new PdfString(da));
            foreach (var w in field.GetWidgets())
            {
                var mk = w.GetPdfObject().GetAsDictionary(PdfName.MK) ?? new PdfDictionary();
                mk.Put(PdfName.CA, new PdfString("4"));
                w.GetPdfObject().Put(PdfName.MK, mk);
                var rect = w.GetRectangle().ToRectangle();
                var normal = new PdfDictionary();
                normal.Put(new PdfName(OnState), CheckAppearance(pdf, rect.GetWidth(), rect.GetHeight(), color));
                normal.Put(new PdfName("Off"), EmptyAppearance(pdf, rect.GetWidth(), rect.GetHeight()));
                var ap = new PdfDictionary();
                ap.Put(PdfName.N, normal);
                w.GetPdfObject().Put(PdfName.AP, ap);
                w.GetPdfObject().Put(PdfName.AS, new PdfName(c.Checked ? OnState : "Off"));
            }
            return field;
        }

        private const string OnState = "Yes";
        private const string ZapfResourceName = "ZaDb";

        private static string F(float v) => v.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);

        /// <summary>✓ をパスで描く（フォントに依存しないのでどのビューアでも同じ形になる）。</summary>
        private static PdfObject CheckAppearance(PdfDocument pdf, float w, float h, Color color)
        {
            var xo = new PdfFormXObject(new Rectangle(0, 0, w, h));
            float s = Math.Min(w, h);
            float ox = (w - s) / 2, oy = (h - s) / 2;
            new PdfCanvas(xo, pdf)
                .SaveState()
                .SetStrokeColor(color)
                .SetLineWidth(Math.Max(0.6f, s * 0.13f))
                .SetLineCapStyle(PdfCanvasConstants.LineCapStyle.ROUND)
                .SetLineJoinStyle(PdfCanvasConstants.LineJoinStyle.ROUND)
                .MoveTo(ox + s * 0.16f, oy + s * 0.53f)
                .LineTo(ox + s * 0.40f, oy + s * 0.22f)
                .LineTo(ox + s * 0.86f, oy + s * 0.84f)
                .Stroke()
                .RestoreState();
            return xo.GetPdfObject();
        }

        private static PdfObject EmptyAppearance(PdfDocument pdf, float w, float h) =>
            new PdfFormXObject(new Rectangle(0, 0, w, h)).GetPdfObject();

        private PdfDictionary _zapf;

        /// <summary>/DR に ZapfDingbats（標準 14 フォント、埋め込み不要）を登録する。</summary>
        private void RegisterZapfDingbats(PdfAcroForm form, PdfDocument pdf)
        {
            if (_zapf == null)
            {
                _zapf = new PdfDictionary();
                _zapf.Put(PdfName.Type, PdfName.Font);
                _zapf.Put(PdfName.Subtype, PdfName.Type1);
                _zapf.Put(PdfName.BaseFont, new PdfName("ZapfDingbats"));
                _zapf.MakeIndirect(pdf);
            }
            FontDictionary(form).Put(new PdfName(ZapfResourceName), _zapf);
        }

        private static PdfFormField CreateRadioGroup(PdfDocument pdf, RadioGroupSpec g, WarningList warnings, out int count)
        {
            count = 0;
            var buttons = g.Buttons.Where(b => CheckPage(pdf, b.PageIndex, g, warnings)).ToList();
            if (buttons.Count == 0) return null;

            // ASCII 以外の Export Value は PDF の名前オブジェクトでは表せないので、
            // 状態名を "0","1",... にして /Opt に本来の値（Unicode 文字列）を持たせる（PDF 1.7 12.7.4.2.4）。
            bool useOpt = buttons.Any(b => b.ExportValue.Any(ch => ch < 0x20 || ch > 0x7E));
            string StateName(int i) => useOpt ? i.ToString(System.Globalization.CultureInfo.InvariantCulture) : buttons[i].ExportValue;

            var builder = new RadioFormFieldBuilder(pdf, g.Name);
            var group = builder.CreateRadioGroup();
            SetUnicodeName(group, g.Name);
            for (int i = 0; i < buttons.Count; i++)
            {
                var b = buttons[i];
                builder.SetPage(b.PageIndex + 1);
                var widget = builder.CreateRadioButton(StateName(i), ToRect(b.Rect));
                widget.SetBorderWidth(0);
                widget.SetBorderColor(null);
                widget.SetBackgroundColor(null);
                group.AddKid(widget);
                count++;
            }

            if (useOpt)
            {
                var opt = new PdfArray();
                foreach (var b in buttons) opt.Add(new PdfString(b.ExportValue, PdfEncodings.UNICODE_BIG));
                group.GetPdfObject().Put(PdfName.Opt, opt);
            }

            int selected = buttons.FindIndex(b => b.ExportValue == g.SelectedValue);
            group.SetValue(selected >= 0 ? StateName(selected) : "Off");

            // Excel のオプションボタンと同様、選択済みのボタンをクリックしても解除しない
            group.SetFieldFlag(PdfButtonFormField.FF_NO_TOGGLE_TO_OFF, true);
            return group;
        }

        /// <summary>日本語のフィールド名が化けないよう /T を UTF-16BE の文字列で書く。</summary>
        private static void SetUnicodeName(PdfFormField field, string name)
        {
            field.GetPdfObject().Put(PdfName.T, new PdfString(name, PdfEncodings.UNICODE_BIG));
        }

        private static void NoBorder(PdfFormField field)
        {
            foreach (var w in field.GetChildFormAnnotations())
            {
                w.SetBorderWidth(0);
                w.SetBorderColor(null);
                w.SetBackgroundColor(null);
            }
        }

        private static PdfDictionary FontDictionary(PdfAcroForm form)
        {
            var dr = form.GetDefaultResources();
            if (dr == null)
            {
                dr = new PdfDictionary();
                form.SetDefaultResources(dr);
            }
            var fonts = dr.GetAsDictionary(PdfName.Font);
            if (fonts == null)
            {
                fonts = new PdfDictionary();
                dr.Put(PdfName.Font, fonts);
            }
            return fonts;
        }

        /// <summary>フォントを AcroForm の /DR に登録する（各 /DA から参照される）。</summary>
        private static void RegisterDefaultResource(PdfAcroForm form, PdfFont font)
        {
            var fonts = FontDictionary(form);
            foreach (var key in fonts.KeySet())
            {
                var existing = fonts.Get(key, false);
                if (existing == font.GetPdfObject() || existing == font.GetPdfObject().GetIndirectReference())
                    return;
            }
            fonts.Put(new PdfName("FJP"), font.GetPdfObject());
        }

        private PdfFont CreateFont(FormPdfSettings settings, WarningList warnings)
        {
            var candidates = new List<(string label, Func<string> path)>();
            if (!string.IsNullOrWhiteSpace(settings.FontFilePath))
                candidates.Add((settings.FontFilePath, () => ResolveFilePath(settings.FontFilePath)));
            candidates.Add((settings.FontName, () => _fontLocator.Find(settings.FontName)));
            foreach (var fb in settings.FallbackFonts ?? new List<string>())
                candidates.Add((fb, () => _fontLocator.Find(fb)));

            bool first = true;
            foreach (var (label, getPath) in candidates)
            {
                if (string.IsNullOrWhiteSpace(label)) continue;
                string path = null;
                try
                {
                    path = getPath();
                    if (path != null)
                    {
                        // サブセット化せずフル埋め込み（入力時に任意の文字を使えるようにする）
                        var font = PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H,
                            PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED);
                        font.SetSubset(false);
                        LastFontDescription = $"{label} ({path})";
                        if (!first)
                            warnings.Add($"フォント「{settings.FontName}」が見つからないため「{label}」を使用しました。");
                        return font;
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add($"フォント「{label}」を読み込めませんでした: {ex.Message}");
                }
                first = false;
            }

            // 指定フォントがどれも無い場合、日本語グリフを持つフォントを探す
            try
            {
                var any = _fontLocator.FindAnyJapanese();
                if (any != null)
                {
                    var font = PdfFontFactory.CreateFont(any, PdfEncodings.IDENTITY_H, PdfFontFactory.EmbeddingStrategy.FORCE_EMBEDDED);
                    font.SetSubset(false);
                    LastFontDescription = any;
                    warnings.Add($"フォント「{settings.FontName}」が見つからないため「{any}」を使用しました。");
                    return font;
                }
            }
            catch (Exception ex)
            {
                Logger.Write("日本語フォントの検索に失敗: " + ex.Message);
            }

            warnings.Add("日本語フォントが見つかりません。テキストフィールドは英数字のみ表示できます。設定でフォントファイルを指定してください。");
            LastFontDescription = "Helvetica";
            return PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA);
        }

        /// <summary>"C:\...\x.ttc" なら ",0" を補う。</summary>
        private static string ResolveFilePath(string path)
        {
            var p = path.Trim().Trim('"');
            int comma = p.LastIndexOf(',');
            string file = comma > 1 && int.TryParse(p.Substring(comma + 1), out _) ? p.Substring(0, comma) : p;
            if (!File.Exists(file)) return null;
            if (file == p && file.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase)) return p + ",0";
            return p;
        }

        /// <summary>校正モード: 各ウィジェットの位置に半透明の赤枠とフィールド名を描く。</summary>
        private static void DrawCalibration(PdfDocument pdf, FormDocumentSpec spec, PdfFont font)
        {
            var boxes = new List<(int page, PdfBox rect, string label)>();
            foreach (var f in spec.Fields)
            {
                switch (f)
                {
                    case TextFieldSpec t: boxes.Add((t.PageIndex, t.Rect, f.FullName)); break;
                    case CheckBoxSpec c: boxes.Add((c.PageIndex, c.Rect, f.FullName)); break;
                    case RadioGroupSpec g:
                        foreach (var b in g.Buttons) boxes.Add((b.PageIndex, b.Rect, f.FullName + "=" + b.ExportValue));
                        break;
                }
            }

            var gs = new PdfExtGState().SetFillOpacity(0.12f).SetStrokeOpacity(0.8f);
            foreach (var pageGroup in boxes.GroupBy(b => b.page))
            {
                if (pageGroup.Key < 0 || pageGroup.Key >= pdf.GetNumberOfPages()) continue;
                var page = pdf.GetPage(pageGroup.Key + 1);
                var canvas = new PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), pdf);
                foreach (var (_, rect, label) in pageGroup)
                {
                    var r = ToRect(rect);
                    canvas.SaveState().SetExtGState(gs).SetLineWidth(0.5f)
                        .SetFillColor(ColorConstants.RED).SetStrokeColor(ColorConstants.RED)
                        .Rectangle(r.GetX(), r.GetY(), r.GetWidth(), r.GetHeight()).FillStroke()
                        .RestoreState();
                    if (font != null)
                    {
                        canvas.SaveState().BeginText().SetFontAndSize(font, 5).SetFillColor(ColorConstants.RED)
                            .MoveText(r.GetX() + 1, r.GetY() + r.GetHeight() + 1).ShowText(label).EndText()
                            .RestoreState();
                    }
                }
            }
        }
    }
}
