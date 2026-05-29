using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace NumDesTools.Advance
{
    /// <summary>
    /// 轻量倒排索引：记录每个 xlsx 文件第2列出现过的数字前缀（取前6位），
    /// 持久化为 JSON，替代 Public.db 的克隆扫描用途。
    /// 文件约几十KB，全量加载到内存后查询为微秒级。
    /// </summary>
    internal static class IdPrefixIndex
    {
        // 取前N位作为前缀粒度——6位覆盖大多数活动ID（73601x、763601x等）
        private const int PrefixLen = 7;
        private const string IndexFileName = "IdPrefixIndex.json";

        private static string IndexPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "NumDesTools",
                "Config",
                IndexFileName
            );

        // ── 持久化格式 ─────────────────────────────────────────────────────────

        private class IndexEntry
        {
            public long Mtime { get; set; }
            public List<string> Prefixes { get; set; } = new();
        }

        private class IndexRoot
        {
            public int Version { get; set; } = 1;
            public Dictionary<string, IndexEntry> Files { get; set; } =
                new(StringComparer.OrdinalIgnoreCase);
        }

        // ── 公开 API ───────────────────────────────────────────────────────────

        /// <summary>
        /// 增量同步 excelDir 下所有 xlsx 到索引（只重扫比索引更新的文件）。
        /// </summary>
        public static void Sync(string excelDir)
        {
            var root = Load();
            var files = Directory.GetFiles(excelDir, "*.xlsx", SearchOption.TopDirectoryOnly);
            var changed = false;

            foreach (var file in files)
            {
                var name = Path.GetFileName(file);
                if (name.StartsWith('#') || name.StartsWith('~'))
                    continue;

                var mtime = new FileInfo(file).LastWriteTimeUtc.Ticks;
                if (root.Files.TryGetValue(name, out var entry) && entry.Mtime == mtime)
                    continue; // 未修改，跳过

                AppServices.App.StatusBar = $"索引更新: {name}";
                var prefixes = ExtractPrefixes(file);
                root.Files[name] = new IndexEntry { Mtime = mtime, Prefixes = prefixes };
                changed = true;
            }

            // 清理已删除的文件
            var existingNames = new HashSet<string>(
                files.Select(Path.GetFileName),
                StringComparer.OrdinalIgnoreCase
            );
            var toRemove = root.Files.Keys.Where(k => !existingNames.Contains(k)).ToList();
            foreach (var k in toRemove)
            {
                root.Files.Remove(k);
                changed = true;
            }

            if (changed)
                Save(root);
        }

        /// <summary>
        /// 查询哪些文件的第2列含有以 activityId 为前缀（或 activityId 以文件前缀开头）的数字。
        /// 返回完整文件名列表（不含路径）。
        /// </summary>
        public static List<string> FindFiles(string activityId)
        {
            var root = Load();
            var result = new List<string>();
            foreach (var (name, entry) in root.Files)
            {
                foreach (var p in entry.Prefixes)
                {
                    // activityId 以 p 开头（p 是短前缀），或 p 以 activityId 开头（activityId 更长）
                    if (
                        activityId.StartsWith(p, StringComparison.Ordinal)
                        || p.StartsWith(activityId, StringComparison.Ordinal)
                    )
                    {
                        result.Add(name);
                        break;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 检查索引是否存在且非空。
        /// </summary>
        public static bool Exists() => File.Exists(IndexPath);

        // ── 内部实现 ───────────────────────────────────────────────────────────

        private static List<string> ExtractPrefixes(string filePath)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                using var pkg = new ExcelPackage(new FileInfo(filePath));
                // 遍历所有 worksheet
                foreach (var sheet in pkg.Workbook.Worksheets)
                {
                    if (sheet.Dimension == null)
                        continue;
                    var endRow = sheet.Dimension.End.Row;
                    // 第2列，从第3行起（跳过类型行和表头行）
                    for (int row = 3; row <= endRow; row++)
                    {
                        var val = sheet.Cells[row, 2].Value;
                        if (val == null)
                            continue;
                        var str = val is double d ? ((long)d).ToString() : val.ToString()!.Trim();
                        if (str.Length < 4)
                            continue;
                        // 必须是纯数字
                        if (!IsAllDigits(str))
                            continue;
                        var prefix = str.Length >= PrefixLen ? str[..PrefixLen] : str;
                        seen.Add(prefix);
                    }
                }
            }
            catch
            {
                // 文件损坏或被锁住，返回空列表，下次会重试
            }
            return seen.ToList();
        }

        private static bool IsAllDigits(string s)
        {
            foreach (var c in s)
                if (!char.IsDigit(c))
                    return false;
            return s.Length > 0;
        }

        private static IndexRoot Load()
        {
            if (!File.Exists(IndexPath))
                return new IndexRoot();
            try
            {
                var json = File.ReadAllText(IndexPath);
                return Newtonsoft.Json.JsonConvert.DeserializeObject<IndexRoot>(json)
                    ?? new IndexRoot();
            }
            catch
            {
                return new IndexRoot();
            }
        }

        private static void Save(IndexRoot root)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(IndexPath)!);
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(
                root,
                Newtonsoft.Json.Formatting.Indented
            );
            File.WriteAllText(IndexPath, json);
        }
    }
}
