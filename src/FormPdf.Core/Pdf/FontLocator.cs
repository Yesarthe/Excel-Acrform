using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using iText.IO.Font;

namespace FormPdf.Core.Pdf
{
    /// <summary>
    /// フォント名（例: "Yu Gothic UI", "MS ゴシック"）からフォントファイルを探す。
    /// TTC の場合は "path,index" 形式（iText の指定方法）を返す。
    /// </summary>
    public sealed class FontLocator
    {
        // よく使う日本語フォントのファイル名ヒント（全フォント走査の前に試す）
        private static readonly Dictionary<string, string[]> Hints = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["yugothicui"] = new[] { "YuGothR.ttc", "YuGothM.ttc" },
            ["yugothic"] = new[] { "YuGothR.ttc", "YuGothM.ttc" },
            ["游ゴシック"] = new[] { "YuGothR.ttc", "YuGothM.ttc" },
            ["msgothic"] = new[] { "msgothic.ttc" },
            ["msゴシック"] = new[] { "msgothic.ttc" },
            ["mspゴシック"] = new[] { "msgothic.ttc" },
            ["msuigothic"] = new[] { "msgothic.ttc" },
            ["meiryo"] = new[] { "meiryo.ttc" },
            ["メイリオ"] = new[] { "meiryo.ttc" },
            ["meiryoui"] = new[] { "meiryo.ttc" },
            ["bizudgothic"] = new[] { "BIZ-UDGothicR.ttc" },
            ["bizudゴシック"] = new[] { "BIZ-UDGothicR.ttc" },
        };

        private readonly List<string> _directories;
        private List<string> _allFiles;

        public FontLocator(IEnumerable<string> extraDirectories = null)
        {
            _directories = new List<string>();
            var windir = Environment.GetEnvironmentVariable("WINDIR");
            if (!string.IsNullOrEmpty(windir)) _directories.Add(Path.Combine(windir, "Fonts"));
            var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (!string.IsNullOrEmpty(local)) _directories.Add(Path.Combine(local, "Microsoft", "Windows", "Fonts"));
            if (extraDirectories != null) _directories.AddRange(extraDirectories);
            _directories = _directories.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>見つかれば iText に渡せるパス（TTC は "path,index"）。無ければ null。</summary>
        public string Find(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName)) return null;
            var key = Normalize(fontName);

            if (Hints.TryGetValue(key, out var hintFiles))
            {
                foreach (var dir in _directories)
                    foreach (var file in hintFiles)
                    {
                        var path = Path.Combine(dir, file);
                        if (File.Exists(path))
                        {
                            var hit = MatchInFile(path, key);
                            if (hit != null) return hit;
                        }
                    }
            }

            foreach (var path in AllFiles())
            {
                var hit = MatchInFile(path, key);
                if (hit != null) return hit;
            }
            return null;
        }

        /// <summary>ひらがな・漢字のグリフを持つフォントを探す（最後の手段）。</summary>
        public string FindAnyJapanese()
        {
            foreach (var name in new[] { "Yu Gothic UI", "Yu Gothic", "MS Gothic", "Meiryo", "BIZ UDGothic", "Noto Sans JP", "Noto Sans CJK JP", "IPAGothic" })
            {
                var hit = Find(name);
                if (hit != null) return hit;
            }
            foreach (var path in AllFiles())
            {
                try
                {
                    if (path.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                    {
                        var ttc = new TrueTypeCollection(path);
                        for (int i = 0; i < ttc.GetTTCSize(); i++)
                            if (HasJapanese(ttc.GetFontByTccIndex(i))) return path + "," + i;
                    }
                    else if (HasJapanese(FontProgramFactory.CreateFont(path)))
                    {
                        return path;
                    }
                }
                catch (Exception)
                {
                    // 読めないフォントは無視
                }
            }
            return null;
        }

        private static bool HasJapanese(FontProgram font) =>
            font.GetGlyph(0x3042) != null && font.GetGlyph(0x6F22) != null; // 「あ」「漢」

        private IEnumerable<string> AllFiles()
        {
            if (_allFiles == null)
            {
                _allFiles = new List<string>();
                foreach (var dir in _directories)
                {
                    try
                    {
                        _allFiles.AddRange(Directory.EnumerateFiles(dir, "*.*", SearchOption.AllDirectories)
                            .Where(f =>
                            {
                                var ext = Path.GetExtension(f).ToLowerInvariant();
                                return ext == ".ttf" || ext == ".ttc" || ext == ".otf";
                            }));
                    }
                    catch (Exception)
                    {
                        // アクセスできないフォルダは無視
                    }
                }
            }
            return _allFiles;
        }

        private static string MatchInFile(string path, string key)
        {
            try
            {
                if (path.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
                {
                    var ttc = new TrueTypeCollection(path);
                    int n = ttc.GetTTCSize();
                    for (int i = 0; i < n; i++)
                    {
                        if (Matches(ttc.GetFontByTccIndex(i), key))
                            return path + "," + i;
                    }
                    return null;
                }
                return Matches(FontProgramFactory.CreateFont(path), key) ? path : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool Matches(FontProgram font, string key)
        {
            var names = font.GetFontNames();
            foreach (var n in Flatten(names.GetFamilyName()).Concat(Flatten(names.GetFamilyName2())))
            {
                if (Normalize(n) == key) return true;
            }
            foreach (var n in Flatten(names.GetFullName()))
            {
                var f = Normalize(n);
                if (f == key || f == key + "regular") return true;
            }
            return false;
        }

        private static IEnumerable<string> Flatten(string[][] names)
        {
            if (names == null) yield break;
            foreach (var row in names)
            {
                if (row != null && row.Length > 0 && !string.IsNullOrEmpty(row[row.Length - 1]))
                    yield return row[row.Length - 1];
            }
        }

        /// <summary>全角→半角（NFKC）、空白除去、小文字化。"ＭＳ ゴシック" と "MS ゴシック" を同一視する。</summary>
        public static string Normalize(string name)
        {
            var s = (name ?? "").Normalize(NormalizationForm.FormKC);
            var sb = new StringBuilder(s.Length);
            foreach (var c in s)
            {
                if (!char.IsWhiteSpace(c) && c != '-' && c != '_') sb.Append(char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }
    }
}
