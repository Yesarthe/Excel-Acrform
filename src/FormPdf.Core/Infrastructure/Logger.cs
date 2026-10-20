using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace FormPdf.Core.Infrastructure
{
    /// <summary>%APPDATA%\FormPdfAddin\log.txt への追記ログ。失敗しても例外は投げない。</summary>
    public static class Logger
    {
        private const long MaxBytes = 2 * 1024 * 1024;
        private static readonly object Sync = new object();

        /// <summary>テスト等でログの出力先を変える場合に設定する。</summary>
        public static string LogPathOverride { get; set; }

        public static string LogPath =>
            LogPathOverride ?? Path.Combine(SettingsStore.AppDataDirectory, "log.txt");

        public static void Write(string message)
        {
            try
            {
                lock (Sync)
                {
                    var path = LogPath;
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    var info = new FileInfo(path);
                    if (info.Exists && info.Length > MaxBytes)
                    {
                        var old = path + ".1";
                        if (File.Exists(old)) File.Delete(old);
                        File.Move(path, old);
                    }
                    File.AppendAllText(path,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}{Environment.NewLine}",
                        new UTF8Encoding(false));
                }
            }
            catch
            {
                // ログ失敗で本処理を止めない
            }
        }

        public static void Write(string message, IEnumerable<string> details)
        {
            var sb = new StringBuilder(message);
            foreach (var d in details)
                sb.Append(Environment.NewLine).Append("    ").Append(d);
            Write(sb.ToString());
        }
    }

    /// <summary>処理を止めずに集める警告。</summary>
    public sealed class WarningList
    {
        private readonly List<string> _items = new List<string>();

        public IReadOnlyList<string> Items => _items;
        public int Count => _items.Count;

        public void Add(string message) => _items.Add(message);

        public void Add(string sheet, string shape, string message) =>
            _items.Add(string.IsNullOrEmpty(shape) ? $"[{sheet}] {message}" : $"[{sheet}] {shape}: {message}");
    }
}
