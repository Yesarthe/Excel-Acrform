using System.Collections.Generic;
using FormPdf.Core.Fields;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Model;

namespace FormPdf.Core.Pdf
{
    /// <summary>
    /// PDF 操作の抽象化。iText 以外（PDFsharp 等）に差し替える場合はこれを実装する。
    /// </summary>
    public interface IPdfFormWriter
    {
        /// <summary>PDF の各ページのサイズ（回転を含む）を返す。</summary>
        IList<PageSizeD> ReadPageSizes(string pdfPath);

        /// <summary>
        /// PDF 内の塗りつぶし（RGB の塗りのみ）を列挙する。自動位置合わせのマーカー検出に使う。
        /// </summary>
        IList<FilledArea> ReadFilledAreas(string pdfPath);

        /// <summary>
        /// PDF 内のすべての描画（塗り・線のパスと画像）の外接矩形を列挙する。
        /// チェックボックスの四角・オプションボタンの丸の位置検出に使う。色は使わない。
        /// </summary>
        IList<FilledArea> ReadDrawnAreas(string pdfPath);

        /// <summary>
        /// sourcePdf（Excel が出力した PDF）に AcroForm フィールドを追加して outputPdf に書き出す。
        /// 致命的でない問題は warnings に追加し、処理を続ける。
        /// </summary>
        /// <returns>作成したフィールド（ウィジェット）の数。</returns>
        int Write(string sourcePdf, string outputPdf, FormDocumentSpec spec, FormPdfSettings settings, WarningList warnings);
    }
}
