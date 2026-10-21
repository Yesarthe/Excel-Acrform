using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using FormPdf.Core.Infrastructure;

namespace FormPdf.AddIn
{
    /// <summary>設定ダイアログ（デザイナーを使わずコードで組み立てる）。</summary>
    internal sealed class SettingsForm : Form
    {
        private readonly ComboBox _fontName = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 260 };
        private readonly TextBox _fontFile = new TextBox { Width = 260 };
        private readonly TextBox _fallbacks = new TextBox { Width = 260 };
        private readonly Button _color = new Button { Width = 120, FlatStyle = FlatStyle.Flat };
        private readonly CheckBox _inheritText = new CheckBox { Text = "テキストボックスの文字を初期値にする", AutoSize = true };
        private readonly CheckBox _inheritChecks = new CheckBox { Text = "チェック／選択状態を初期値にする", AutoSize = true };
        private readonly CheckBox _scaleFont = new CheckBox { Text = "フォントサイズを印刷倍率に合わせる", AutoSize = true };
        private readonly CheckBox _openAfter = new CheckBox { Text = "出力後に PDF を開く", AutoSize = true };
        private readonly CheckBox _autoCalibrate = new CheckBox { Text = "自動位置合わせ（Excel で 2 回出力して実際の位置を測る）", AutoSize = true };
        private readonly CheckBox _calibration = new CheckBox { Text = "校正モード（フィールド位置に赤枠と名前を描く）", AutoSize = true };
        private readonly CheckBox _detectButtons = new CheckBox { Text = "四角・丸の位置を Excel の印字から検出する", AutoSize = true };
        private readonly NumericUpDown _offsetX = Number(-200, 200, 1, 0.5m);
        private readonly NumericUpDown _offsetY = Number(-200, 200, 1, 0.5m);
        private readonly NumericUpDown _scaleFix = Number(-20, 20, 2, 0.1m);
        private readonly NumericUpDown _checkSize = Number(0, 100, 1, 0.5m);
        private readonly NumericUpDown _optionSize = Number(0, 100, 1, 0.5m);
        private readonly NumericUpDown _buttonInset = Number(-20, 50, 1, 0.5m);
        private Color _checkColor;
        private readonly FormPdfSettings _original;

        public FormPdfSettings Result { get; private set; }

        public SettingsForm(FormPdfSettings settings)
        {
            _original = settings;
            Text = "フォームPDF 設定";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(10);

            _fontName.Items.AddRange(new object[] { "Yu Gothic UI", "游ゴシック", "MS ゴシック", "MS UI Gothic", "メイリオ", "BIZ UDゴシック" });

            var table = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            Section(table, "テキストフィールド");
            Row(table, "フォント名", _fontName);
            var browse = new Button { Text = "参照…", AutoSize = true };
            browse.Click += (s, e) => BrowseFont();
            Row(table, "フォントファイル（任意）", Flow(_fontFile, browse));
            Row(table, "代替フォント（; 区切り）", _fallbacks);
            Row(table, "", _inheritText);
            Row(table, "", _scaleFont);

            Section(table, "チェックボックス・オプションボタン");
            _color.Click += (s, e) => PickColor();
            Row(table, "チェック記号の色", _color);
            Row(table, "", _inheritChecks);
            Row(table, "", _detectButtons);
            Row(table, "検出できないときの四角（pt、0=高さ）", _checkSize);
            Row(table, "検出できないときの丸（pt、0=高さ）", _optionSize);
            Row(table, "四角・丸の左端位置（pt）", _buttonInset);

            Section(table, "位置の補正（校正）");
            Row(table, "", _autoCalibrate);
            Row(table, "", _calibration);
            Row(table, "X 補正（pt、右が＋）", _offsetX);
            Row(table, "Y 補正（pt、下が＋）", _offsetY);
            Row(table, "倍率補正（%）", _scaleFix);

            Section(table, "その他");
            Row(table, "", _openAfter);

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 90 };
            var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, Width = 90 };
            var reset = new Button { Text = "既定に戻す", AutoSize = true };
            var folder = new Button { Text = "設定フォルダを開く", AutoSize = true };
            ok.Click += (s, e) => Result = Collect();
            reset.Click += (s, e) => LoadValues(new FormPdfSettings { Scope = _original.Scope });
            folder.Click += (s, e) =>
            {
                System.IO.Directory.CreateDirectory(SettingsStore.AppDataDirectory);
                Dialogs.OpenPath(SettingsStore.AppDataDirectory);
            };
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 12, 0, 0) };
            buttons.Controls.AddRange(new Control[] { cancel, ok, reset, folder });
            table.Controls.Add(buttons, 0, table.RowCount);
            table.SetColumnSpan(buttons, 2);

            Controls.Add(table);
            AcceptButton = ok;
            CancelButton = cancel;

            LoadValues(settings);
        }

        private static NumericUpDown Number(decimal min, decimal max, int decimals, decimal step) =>
            new NumericUpDown { Minimum = min, Maximum = max, DecimalPlaces = decimals, Increment = step, Width = 90, TextAlign = HorizontalAlignment.Right };

        private static Control Flow(params Control[] controls)
        {
            var p = new FlowLayoutPanel { AutoSize = true, Margin = Padding.Empty, WrapContents = false };
            p.Controls.AddRange(controls);
            return p;
        }

        private static void Section(TableLayoutPanel table, string title)
        {
            var label = new Label
            {
                Text = title,
                AutoSize = true,
                Font = new Font(SystemFonts.MessageBoxFont, FontStyle.Bold),
                Margin = new Padding(0, table.RowCount == 0 ? 0 : 10, 0, 4),
            };
            table.Controls.Add(label, 0, table.RowCount);
            table.SetColumnSpan(label, 2);
            table.RowCount++;
        }

        private static void Row(TableLayoutPanel table, string caption, Control control)
        {
            var label = new Label { Text = caption, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(12, 6, 8, 3) };
            control.Anchor = AnchorStyles.Left;
            table.Controls.Add(label, 0, table.RowCount);
            table.Controls.Add(control, 1, table.RowCount);
            table.RowCount++;
        }

        private void LoadValues(FormPdfSettings s)
        {
            _fontName.Text = s.FontName;
            _fontFile.Text = s.FontFilePath;
            _fallbacks.Text = string.Join("; ", s.FallbackFonts ?? new System.Collections.Generic.List<string>());
            var (r, g, b) = s.GetCheckRgb();
            SetColor(Color.FromArgb(r, g, b));
            _inheritText.Checked = s.InheritTextValues;
            _inheritChecks.Checked = s.InheritCheckStates;
            _scaleFont.Checked = s.ScaleFontWithPage;
            _openAfter.Checked = s.OpenAfterExport;
            _autoCalibrate.Checked = s.AutoCalibrate;
            _calibration.Checked = s.CalibrationMode;
            _offsetX.Value = Clamp(_offsetX, s.OffsetX);
            _offsetY.Value = Clamp(_offsetY, s.OffsetY);
            _scaleFix.Value = Clamp(_scaleFix, s.ScaleCorrectionPercent);
            _detectButtons.Checked = s.DetectButtonFrames;
            _checkSize.Value = Clamp(_checkSize, s.CheckBoxSize);
            _optionSize.Value = Clamp(_optionSize, s.OptionButtonSize);
            _buttonInset.Value = Clamp(_buttonInset, s.ButtonLeftInset);
        }

        private static decimal Clamp(NumericUpDown n, double v) =>
            Math.Max(n.Minimum, Math.Min(n.Maximum, (decimal)v));

        private FormPdfSettings Collect()
        {
            var s = _original.Clone();
            s.FontName = _fontName.Text.Trim();
            s.FontFilePath = _fontFile.Text.Trim();
            s.FallbackFonts = _fallbacks.Text.Split(new[] { ';', '；' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
            s.CheckColor = $"#{_checkColor.R:X2}{_checkColor.G:X2}{_checkColor.B:X2}";
            s.InheritTextValues = _inheritText.Checked;
            s.InheritCheckStates = _inheritChecks.Checked;
            s.ScaleFontWithPage = _scaleFont.Checked;
            s.OpenAfterExport = _openAfter.Checked;
            s.AutoCalibrate = _autoCalibrate.Checked;
            s.CalibrationMode = _calibration.Checked;
            s.OffsetX = (double)_offsetX.Value;
            s.OffsetY = (double)_offsetY.Value;
            s.ScaleCorrectionPercent = (double)_scaleFix.Value;
            s.DetectButtonFrames = _detectButtons.Checked;
            s.CheckBoxSize = (double)_checkSize.Value;
            s.OptionButtonSize = (double)_optionSize.Value;
            s.ButtonLeftInset = (double)_buttonInset.Value;
            return s;
        }

        private void SetColor(Color c)
        {
            _checkColor = c;
            _color.BackColor = c;
            _color.ForeColor = c.GetBrightness() < 0.5 ? Color.White : Color.Black;
            _color.Text = string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", c.R, c.G, c.B);
        }

        private void PickColor()
        {
            using (var dlg = new ColorDialog { Color = _checkColor, FullOpen = true })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK) SetColor(dlg.Color);
            }
        }

        private void BrowseFont()
        {
            using (var dlg = new OpenFileDialog
            {
                Title = "フォントファイルを選択",
                Filter = "フォント (*.ttf;*.ttc;*.otf)|*.ttf;*.ttc;*.otf",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
            })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK) _fontFile.Text = dlg.FileName;
            }
        }
    }
}
