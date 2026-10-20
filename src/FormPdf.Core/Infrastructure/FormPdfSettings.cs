using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;

namespace FormPdf.Core.Infrastructure
{
    public enum ExportScope
    {
        ActiveSheet = 0,
        SelectedSheets = 1,
        Workbook = 2,
    }

    /// <summary>%APPDATA%\FormPdfAddin\settings.json に保存する設定。</summary>
    [DataContract]
    public sealed class FormPdfSettings
    {
        /// <summary>テキストフィールドのフォント（ファミリー名）。</summary>
        [DataMember(Order = 1)] public string FontName { get; set; } = "Yu Gothic UI";
        /// <summary>フォントファイルを直接指定する場合のパス（.ttf/.ttc/.otf）。TTC は "path,index" も可。空ならフォント名で検索。</summary>
        [DataMember(Order = 2)] public string FontFilePath { get; set; } = "";
        /// <summary>FontName が見つからないときに順に試すフォント。</summary>
        [DataMember(Order = 3)] public List<string> FallbackFonts { get; set; } = new List<string> { "MS ゴシック", "MS UI Gothic" };

        /// <summary>チェック記号の色（#RRGGBB）。</summary>
        [DataMember(Order = 10)] public string CheckColor { get; set; } = "#FF0000";

        /// <summary>テキストボックスの文字列を初期値として引き継ぐ。false なら空欄。</summary>
        [DataMember(Order = 20)] public bool InheritTextValues { get; set; } = true;
        /// <summary>チェックボックス／オプションボタンの状態を初期値として引き継ぐ。</summary>
        [DataMember(Order = 21)] public bool InheritCheckStates { get; set; } = true;
        /// <summary>印刷時の縮小率に合わせてフォントサイズも縮小する（見た目を Excel の印刷結果に合わせる）。</summary>
        [DataMember(Order = 22)] public bool ScaleFontWithPage { get; set; } = true;

        /// <summary>出力後に PDF を既定アプリで開く。</summary>
        [DataMember(Order = 30)] public bool OpenAfterExport { get; set; } = true;
        [DataMember(Order = 31)] public ExportScope Scope { get; set; } = ExportScope.ActiveSheet;

        /// <summary>校正モード: 各フィールド位置に赤枠と名前を描く。</summary>
        [DataMember(Order = 40)] public bool CalibrationMode { get; set; }
        /// <summary>全体の X 補正（pt、右が正）。</summary>
        [DataMember(Order = 41)] public double OffsetX { get; set; }
        /// <summary>全体の Y 補正（pt、下が正）。</summary>
        [DataMember(Order = 42)] public double OffsetY { get; set; }
        /// <summary>倍率の補正（%）。例: 0.5 なら倍率を 1.005 倍にする。</summary>
        [DataMember(Order = 43)] public double ScaleCorrectionPercent { get; set; }

        /// <summary>チェックボックスの四角部分のサイズ（pt、Excel 上の 100% 時）。0 なら図形の高さを使う。</summary>
        [DataMember(Order = 50)] public double CheckBoxSize { get; set; } = 10;
        /// <summary>オプションボタンの丸部分のサイズ（pt、Excel 上の 100% 時）。0 なら図形の高さを使う。</summary>
        [DataMember(Order = 51)] public double OptionButtonSize { get; set; } = 10;
        /// <summary>四角・丸の左端の位置（図形左端からの距離、pt）。</summary>
        [DataMember(Order = 52)] public double ButtonLeftInset { get; set; } = 0;

        [OnDeserializing]
        private void OnDeserializing(StreamingContext _)
        {
            // DataContract はコンストラクタを通らないので既定値をここで入れる
            var d = new FormPdfSettings();
            FontName = d.FontName;
            FontFilePath = d.FontFilePath;
            FallbackFonts = d.FallbackFonts;
            CheckColor = d.CheckColor;
            InheritTextValues = d.InheritTextValues;
            InheritCheckStates = d.InheritCheckStates;
            ScaleFontWithPage = d.ScaleFontWithPage;
            OpenAfterExport = d.OpenAfterExport;
            Scope = d.Scope;
            CheckBoxSize = d.CheckBoxSize;
            OptionButtonSize = d.OptionButtonSize;
            ButtonLeftInset = d.ButtonLeftInset;
        }

        public FormPdfSettings Clone()
        {
            var json = ToJson();
            return FromJson(json);
        }

        /// <summary>#RRGGBB を (r,g,b) に変換する。不正なら赤。</summary>
        public (int R, int G, int B) GetCheckRgb()
        {
            var s = (CheckColor ?? "").Trim().TrimStart('#');
            if (s.Length == 6 &&
                int.TryParse(s.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                int.TryParse(s.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                int.TryParse(s.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
            {
                return (r, g, b);
            }
            return (255, 0, 0);
        }

        public string ToJson()
        {
            var serializer = new DataContractJsonSerializer(typeof(FormPdfSettings),
                new DataContractJsonSerializerSettings { UseSimpleDictionaryFormat = true });
            using (var ms = new MemoryStream())
            {
                using (var writer = JsonReaderWriterFactory.CreateJsonWriter(ms, Encoding.UTF8, false, true, "  "))
                {
                    serializer.WriteObject(writer, this);
                }
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public static FormPdfSettings FromJson(string json)
        {
            var serializer = new DataContractJsonSerializer(typeof(FormPdfSettings));
            using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                var s = (FormPdfSettings)serializer.ReadObject(ms);
                if (s.FallbackFonts == null) s.FallbackFonts = new List<string>();
                return s;
            }
        }
    }

    public static class SettingsStore
    {
        public static string AppDataDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "FormPdfAddin");

        public static string SettingsPath => Path.Combine(AppDataDirectory, "settings.json");

        public static FormPdfSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                    return FormPdfSettings.FromJson(File.ReadAllText(SettingsPath, Encoding.UTF8));
            }
            catch (Exception ex)
            {
                Logger.Write("設定ファイルを読めなかったため既定値を使います: " + ex.Message);
            }
            return new FormPdfSettings();
        }

        public static void Save(FormPdfSettings settings)
        {
            Directory.CreateDirectory(AppDataDirectory);
            File.WriteAllText(SettingsPath, settings.ToJson(), new UTF8Encoding(false));
        }
    }
}
