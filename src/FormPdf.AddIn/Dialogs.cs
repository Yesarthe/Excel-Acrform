using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using FormPdf.Core.Infrastructure;
using XL = Microsoft.Office.Interop.Excel;

namespace FormPdf.AddIn
{
    /// <summary>Excel のメインウィンドウをダイアログの親にする。</summary>
    internal sealed class ExcelWindow : IWin32Window
    {
        public ExcelWindow(XL.Application app)
        {
            try { Handle = new IntPtr(app.Hwnd); } catch { Handle = IntPtr.Zero; }
        }

        public IntPtr Handle { get; }
    }

    internal static class Dialogs
    {
        public const string Title = "フォームPDF";

        public static void Info(IWin32Window owner, string message) =>
            MessageBox.Show(owner, message, Title, MessageBoxButtons.OK, MessageBoxIcon.Information);

        public static void Error(IWin32Window owner, string message) =>
            MessageBox.Show(owner, message, Title, MessageBoxButtons.OK, MessageBoxIcon.Error);

        public static DialogResult YesNo(IWin32Window owner, string message) =>
            MessageBox.Show(owner, message, Title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);

        /// <summary>警告一覧（スクロール可能なテキスト）を表示する。</summary>
        public static void Warnings(IWin32Window owner, string header, IEnumerable<string> items)
        {
            using (var form = new Form
            {
                Text = Title,
                StartPosition = FormStartPosition.CenterParent,
                Size = new Size(640, 420),
                MinimizeBox = false,
                MaximizeBox = false,
                ShowIcon = false,
                ShowInTaskbar = false,
                Font = SystemFonts.MessageBoxFont,
                AutoScaleMode = AutoScaleMode.Dpi,
            })
            {
                var label = new Label { Text = header, Dock = DockStyle.Top, AutoSize = false, Height = 44, Padding = new Padding(8, 8, 8, 0) };
                var box = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false,
                    Dock = DockStyle.Fill,
                    Text = string.Join(Environment.NewLine, items.Select(i => "・" + i)),
                };
                var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 40, Padding = new Padding(4) };
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
                var log = new Button { Text = "ログを開く", Width = 110 };
                log.Click += (s, e) => OpenPath(Logger.LogPath);
                bottom.Controls.Add(ok);
                bottom.Controls.Add(log);
                form.Controls.Add(box);
                form.Controls.Add(label);
                form.Controls.Add(bottom);
                form.AcceptButton = ok;
                form.ShowDialog(owner);
            }
        }

        public static void OpenPath(string path)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Title);
            }
        }
    }
}
