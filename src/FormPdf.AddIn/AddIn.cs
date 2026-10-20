using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using FormPdf.Core.Infrastructure;
using FormPdf.Core.Pdf;
using XL = Microsoft.Office.Interop.Excel;

namespace FormPdf.AddIn
{
    public sealed class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            Logger.Write("アドインを読み込みました。");
        }

        public void AutoClose()
        {
        }
    }

    /// <summary>リボンの「フォームPDF」タブ。</summary>
    [ComVisible(true)]
    public sealed class FormPdfRibbon : ExcelRibbon
    {
        private FormPdfSettings _settings;

        private FormPdfSettings Settings => _settings ?? (_settings = SettingsStore.Load());

        public override string GetCustomUI(string ribbonId) => @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='FormPdfTab' label='フォームPDF'>
        <group id='FormPdfGroup' label='フォームPDF'>
          <button id='FormPdfExport' label='フォームPDF出力' size='large' imageMso='FileSaveAsPdfOrXps'
                  screentip='フォームPDF出力'
                  supertip='テキストボックス・チェックボックス・オプションボタンを入力可能なフィールドにした PDF を出力します。'
                  onAction='OnExport' />
          <dropDown id='FormPdfScope' label='出力範囲' sizeString='選択中のシートWW'
                    getSelectedItemIndex='GetScopeIndex' onAction='OnScopeChanged'>
            <item id='ScopeActive' label='アクティブシート' />
            <item id='ScopeSelected' label='選択中のシート' />
            <item id='ScopeWorkbook' label='ブック全体' />
          </dropDown>
          <button id='FormPdfSettings' label='設定' size='large' imageMso='PropertySheet'
                  screentip='設定' supertip='フォント、チェックの色、初期値、校正モードなどを設定します。'
                  onAction='OnSettings' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public void OnExport(IRibbonControl control)
        {
            var app = (XL.Application)ExcelDnaUtil.Application;
            try
            {
                _settings = SettingsStore.Load();
                new ExportService(app, new ITextPdfFormWriter()).Run(_settings);
            }
            catch (Exception ex)
            {
                Logger.Write("エラー: " + ex);
                Dialogs.Error(new ExcelWindow(app), "PDF の出力中にエラーが発生しました。\n\n" + ex.Message + "\n\n詳細: " + Logger.LogPath);
            }
            finally
            {
                try { app.StatusBar = false; } catch { }
            }
        }

        public int GetScopeIndex(IRibbonControl control) => (int)Settings.Scope;

        public void OnScopeChanged(IRibbonControl control, string selectedId, int selectedIndex)
        {
            var s = SettingsStore.Load();
            s.Scope = (ExportScope)Math.Max(0, Math.Min(2, selectedIndex));
            _settings = s;
            TrySave(s);
        }

        public void OnSettings(IRibbonControl control)
        {
            var app = (XL.Application)ExcelDnaUtil.Application;
            using (var form = new SettingsForm(SettingsStore.Load()))
            {
                if (form.ShowDialog(new ExcelWindow(app)) == DialogResult.OK)
                {
                    _settings = form.Result;
                    TrySave(_settings);
                }
            }
        }

        private static void TrySave(FormPdfSettings s)
        {
            try
            {
                SettingsStore.Save(s);
            }
            catch (Exception ex)
            {
                Logger.Write("設定を保存できませんでした: " + ex.Message);
                MessageBox.Show("設定を保存できませんでした。\n" + ex.Message, Dialogs.Title);
            }
        }
    }
}
