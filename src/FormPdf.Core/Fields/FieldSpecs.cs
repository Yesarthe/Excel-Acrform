using System.Collections.Generic;
using FormPdf.Core.Model;

namespace FormPdf.Core.Fields
{
    /// <summary>PDF に作るフィールドの共通部分。</summary>
    public abstract class FieldSpec
    {
        /// <summary>親フィールド名（複数シート出力時のシート名）。単一シートなら null。</summary>
        public string ParentName { get; set; }
        /// <summary>部分フィールド名（ピリオドを含まない）。</summary>
        public string Name { get; set; }
        /// <summary>警告表示用: 元のシート名・図形名。</summary>
        public string SourceSheet { get; set; }
        public string SourceShape { get; set; }

        public string FullName => string.IsNullOrEmpty(ParentName) ? Name : ParentName + "." + Name;
    }

    public sealed class TextFieldSpec : FieldSpec
    {
        /// <summary>0 始まりの PDF ページ番号。</summary>
        public int PageIndex { get; set; }
        public PdfBox Rect { get; set; }
        public string Value { get; set; } = "";
        public double FontSize { get; set; }
        public HAlign Alignment { get; set; }
        public bool Multiline { get; set; }
    }

    public sealed class CheckBoxSpec : FieldSpec
    {
        public int PageIndex { get; set; }
        public PdfBox Rect { get; set; }
        public bool Checked { get; set; }
    }

    public sealed class RadioButtonSpec
    {
        public string ExportValue { get; set; }
        public int PageIndex { get; set; }
        public PdfBox Rect { get; set; }
        public string SourceShape { get; set; }
    }

    public sealed class RadioGroupSpec : FieldSpec
    {
        public List<RadioButtonSpec> Buttons { get; } = new List<RadioButtonSpec>();
        /// <summary>選択状態のボタンの ExportValue。無ければ null。</summary>
        public string SelectedValue { get; set; }
    }

    /// <summary>1 つの PDF に作るフィールド一式。</summary>
    public sealed class FormDocumentSpec
    {
        public List<FieldSpec> Fields { get; } = new List<FieldSpec>();
        public int Count => Fields.Count;
    }
}
