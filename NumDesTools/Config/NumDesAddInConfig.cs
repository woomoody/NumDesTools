using System.Text.Json;

namespace NumDesTools.Config
{
    public class NumDesToolsConfig
    {
        #region 默认值

        private readonly List<string> _defaultNormaKeyList =
        [
            ",,",
            "[,",
            ",]",
            "{,",
            ",}",
            "，，",
            "[，",
            "，]",
            "{，",
            "，}"
        ];
        private readonly List<string> _defaultSpecialKeyList = ["][", "}{"];
        private readonly List<CoupleKey> _defaultCoupleKeyList =
        [
            new CoupleKey("[", "]"),
            new CoupleKey("{", "}")
        ];

        #endregion

        private ConfigData _configData;

        private readonly string _filePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NumDesToolsConfig.json"
        );

        public NumDesToolsConfig()
        {
            LoadOrCreateConfig();
        }

        public List<string> NormaKeyList => _configData.NormaKeyList;
        public List<string> SpecialKeyList => _configData.SpecialKeyList;
        public List<CoupleKey> CoupleKeyList => _configData.CoupleKeyList;

        private void LoadOrCreateConfig()
        {
            if (File.Exists(_filePath))
            {
                var json = File.ReadAllText(_filePath);
                _configData = JsonSerializer.Deserialize<ConfigData>(json);

                if (_configData.Equals(default(ConfigData)))
                {
                    _configData = CreateDefaultConfig();
                }
                else
                {
                    MergeWithDefaults();
                }
            }
            else
            {
                _configData = CreateDefaultConfig();
                SaveConfig();
            }
        }

        private void MergeWithDefaults()
        {
            //合并一般字符
            _configData.NormaKeyList = MergeLists(_configData.NormaKeyList, _defaultNormaKeyList);
            //合并特殊字符
            _configData.SpecialKeyList = MergeLists(
                _configData.SpecialKeyList,
                _defaultSpecialKeyList
            );
            //合并成对字符
            _configData.CoupleKeyList = MergeLists(
                _configData.CoupleKeyList,
                _defaultCoupleKeyList
            );
        }

        private List<T> MergeLists<T>(List<T> original, List<T> defaults)
        {
            var result = new HashSet<T>(original);
            result.UnionWith(defaults);
            return [.. result];
        }

        private ConfigData CreateDefaultConfig()
        {
            return new ConfigData
            {
                NormaKeyList = [.. _defaultNormaKeyList],
                SpecialKeyList = [.. _defaultSpecialKeyList],
                CoupleKeyList = [.. _defaultCoupleKeyList]
            };
        }

        private void SaveConfig()
        {
            var json = JsonSerializer.Serialize(
                _configData,
                new JsonSerializerOptions { WriteIndented = true }
            );
            File.WriteAllText(_filePath, json);
        }

        private struct ConfigData
        {
            public List<string> NormaKeyList { get; set; }
            public List<string> SpecialKeyList { get; set; }
            public List<CoupleKey> CoupleKeyList { get; set; }
        }

        public struct CoupleKey(string left, string right)
        {
            public string Left { get; set; } = left;
            public string Right { get; set; } = right;

            public void Deconstruct(out string left, out string right)
            {
                left = Left;
                right = Right;
            }
        }
    }
}
