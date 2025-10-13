using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace NumDesTools;

// Pure logic helpers extracted from LteData to enable unit testing and reduce COM/UI coupling
internal static class LteCore
{
    private static readonly Regex WildcardRegex = new("#(.*?)#", RegexOptions.Compiled);

    public static string AnalyzeWildcard(
        string cellModelValue,
        Dictionary<string, string> exportWildcardData,
        Dictionary<string, string> exportWildcardDyData,
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        Dictionary<string, List<string>> baseData,
        string id,
        string itemId
    )
    {
        var cellRealValue = cellModelValue ?? string.Empty;
        var wildcardValuePattern = "#";
        List<string> idList = baseData.ContainsKey(id) ? baseData[id] : new List<string>();

        MatchCollection matches = WildcardRegex.Matches(cellModelValue ?? string.Empty);

        foreach (Match match in matches)
        {
            var wildcard = match.Groups[1].Value;
            if (!exportWildcardData.TryGetValue(wildcard, out var wildcardValue))
            {
                continue;
            }

            var wildcardValueSplit = Regex.Split(wildcardValue, wildcardValuePattern);
            string funName = wildcardValueSplit.ElementAtOrDefault(0) ?? "";
            string funDepends = wildcardValueSplit.ElementAtOrDefault(1) ?? "物品编号";
            string funDy1 = wildcardValueSplit.ElementAtOrDefault(2) ?? "";
            string funDy2 = wildcardValueSplit.ElementAtOrDefault(3) ?? "";
            string funDy3 = wildcardValueSplit.ElementAtOrDefault(4) ?? "";
            string funDy4 = wildcardValueSplit.ElementAtOrDefault(5) ?? "";
            string funDy5 = wildcardValueSplit.ElementAtOrDefault(6) ?? "";

            try
            {
                string fixWildcardValue = funName switch
                {
                    //根据动态或静态值计算值
                    "Left" => Left(exportWildcardDyData, funDepends, funDy1),
                    "Right" => Right(exportWildcardDyData, funDepends, funDy1),
                    "Set" => Set(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "SetDic"
                        => SetDic(
                            exportWildcardDyData,
                            strDictionary,
                            wildcard,
                            funDepends,
                            funDy1,
                            funDy2,
                            funDy3,
                            idList
                        ),
                    "Mer" => Mer(exportWildcardDyData, funDepends, itemId, funDy1),
                    "MerB"
                        => MerB(exportWildcardDyData, funDepends, itemId, funDy1, funDy2, funDy3),
                    "MerTry"
                        => MerTry(exportWildcardDyData, funDepends, funDy1, funDy2, funDy3, idList),
                    "Ads" => Ads(exportWildcardDyData, funDepends, funDy1, idList),
                    "Arr" => Arr(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "Get" => Get(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "GetDic"
                        => GetDic(
                            strDictionary,
                            exportWildcardDyData,
                            funDepends,
                            funDy1,
                            funDy2,
                            funDy3
                        ),
                    "GetDicKey" => GetDicKey(funDepends),
                    "SplitArr" => SplitArr(exportWildcardDyData, funDepends, funDy1, funDy2),
                    "CollectRow"
                        => CollectRow(
                            exportWildcardDyData,
                            funDepends,
                            funDy1,
                            funDy2,
                            funDy3,
                            funDy4,
                            funDy5,
                            baseData,
                            id
                        ),
                    //获取动态值
                    "Var" => exportWildcardDyData.ContainsKey(wildcard) ? exportWildcardDyData[wildcard] : string.Empty,

                    //获取静态值
                    _ => exportWildcardData.ContainsKey(wildcard) ? exportWildcardData[wildcard] : string.Empty
                };

                cellRealValue = cellRealValue.Replace($"#{wildcard}#", fixWildcardValue);
            }
            catch (FormatException)
            {
                Debug.Print($"通配符解析错误: {wildcard} | 值: {exportWildcardDyData.GetValueOrDefault(wildcard)}");
                return string.Empty;
            }
        }

        return cellRealValue;
    }

    public static string Left(Dictionary<string, string> exportWildcardDyData, string funDepends, string funDy1)
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsValue)) return string.Empty;
        var maxCount = Math.Min(dependsValue.Length, int.Parse(funDy1));
        return dependsValue.Substring(0, maxCount);
    }

    public static string Right(Dictionary<string, string> exportWildcardDyData, string funDepends, string funDy1)
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsValue)) return string.Empty;
        var maxCount = Math.Min(dependsValue.Length, int.Parse(funDy1));
        return dependsValue.Substring(dependsValue.Length - maxCount, int.Parse(funDy1));
    }

    public static string Set(Dictionary<string, string> exportWildcardDyData, string funDepends, string funDy1, string funDy2)
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "00" : funDy2;
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsValue)) return string.Empty;
        return dependsValue.Substring(0, dependsValue.Length - int.Parse(funDy1)) + funDy2;
    }

    public static string SetDic(
        Dictionary<string, string> exportWildcardDyData,
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        string wildcard,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3,
        List<string> idList
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "2" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "00" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "链类最大值" : funDy3;

        string fixWildcardValue = Set(exportWildcardDyData, funDepends, funDy1, funDy2);
        InitializeDictionary(strDictionary, wildcard, fixWildcardValue);

        if (!exportWildcardDyData.TryGetValue(funDy3, out var maxLink))
            maxLink = string.Empty;

        if (maxLink != "")
        {
            var linkList = new List<string>();
            for (int i = 0; i < int.Parse(maxLink); i++)
            {
                var tempId = (long.Parse(fixWildcardValue) + i + 1).ToString();
                if (idList.Contains(tempId))
                {
                    linkList.Add(tempId);
                }
            }
            strDictionary[wildcard][fixWildcardValue] = linkList;
        }
        return fixWildcardValue;
    }

    public static string Mer(Dictionary<string, string> exportWildcardDyData, string funDepends, string itemId, string funDy1)
    {
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsValue)) return string.Empty;

        if (long.TryParse(dependsValue, out long value))
        {
            return (value + int.Parse(funDy1)).ToString();
        }

        Debug.Print($"Mer: 无法将 '{dependsValue}' 解析为 long 类型。");
        return dependsValue;
    }

    public static string MerB(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string itemId,
        string funDy1,
        string funDy2,
        string funDy3
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "3" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "10" : funDy3;
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsValue)) return "0";

        var baseValue = dependsValue.Substring(dependsValue.Length - 1, 1);

        if (int.TryParse(baseValue, out int baseValueTry))
        {
            if (int.TryParse(funDy1, out int funDy1Try))
            {
                if (int.TryParse(funDy2, out int funDy2Try))
                {
                    if (long.TryParse(dependsValue, out long exValue))
                    {
                        if (int.TryParse(funDy3, out int funDy3Try))
                        {
                            if (baseValueTry + funDy1Try <= funDy2Try)
                            {
                                return (exValue + funDy1Try).ToString();
                            }
                            else
                            {
                                return (exValue + funDy1Try + funDy3Try).ToString();
                            }
                        }
                    }
                }
            }
        }

        Debug.Print($"MerB: 无法解析参数，返回 0");
        return "0";
    }

    public static string MerTry(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3,
        List<string> idList
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "3" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "10" : funDy3;
        string merB = MerB(exportWildcardDyData, funDepends, string.Empty, funDy1, funDy2, funDy3);
        var mer = !idList.Contains(merB) ? Mer(exportWildcardDyData, funDepends, string.Empty, funDy1) : merB;
        return mer;
    }

    public static string Ads(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        List<string> idList
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "链类最大值" : funDy1;
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsValue)) return string.Empty;

        string rootNum = dependsValue.Substring(0, dependsValue.Length - 2) + "00";
        int baseValue = int.Parse(dependsValue.Substring(dependsValue.Length - 1, 1));
        int baseMax = 0;
        if (funDy1 != "链类最大值")
        {
            return string.Empty;
        }
        try
        {
            baseMax = int.Parse(exportWildcardDyData.GetValueOrDefault(funDy1, "0"));
        }
        catch (Exception e)
        {
            Debug.Print($"{rootNum}##{funDy1}可能为空 {e.Message}");
        }
        if (baseMax == 0)
        {
            Debug.Print($"{rootNum}物品应该不属于链");
        }
        var loopNum = LoopNumber(baseValue, baseMax);
        var resultSb = new StringBuilder();
        foreach (var num in loopNum)
        {
            var digNum = (long.Parse(rootNum) + num).ToString();
            if (idList.Contains(digNum))
            {
                resultSb.Append(digNum).Append(',');
            }
        }

        var result = resultSb.Length > 0 ? resultSb.ToString(0, resultSb.Length - 1) : string.Empty;
        return result;
    }

    public static string Arr(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "消耗量组" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "" : funDy2;

        var funDy1Value = exportWildcardDyData.GetValueOrDefault(funDy1, "");
        var funDependsValue = exportWildcardDyData.GetValueOrDefault(funDepends, "");

        var funDy1ValueSplit = Regex.Split(funDy1Value, ",");
        var funDependsValueSplit = Regex.Split(funDependsValue, ",");

        var sb = new StringBuilder();
        if (funDy1ValueSplit.Length == funDependsValueSplit.Length)
        {
            for (int i = 0; i < funDy1ValueSplit.Length; i++)
            {
                string temp;
                if (funDy2 != "")
                {
                    temp = $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]},{funDependsValueSplit[i]}]";
                }
                else
                {
                    temp = $"[{funDependsValueSplit[i]},{funDy1ValueSplit[i]}]";
                }
                sb.Append(temp).Append(',');
            }
            if (sb.Length > 0) sb.Length--; // remove trailing comma
        }
        return sb.ToString();
    }

    public static string Get(Dictionary<string, string> exportWildcardDyData, string funDepends, string funDy1, string funDy2)
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "," : funDy2;
        var dependsValue = exportWildcardDyData.GetValueOrDefault(funDepends, "");
        var dependsValueSplit = Regex.Split(dependsValue, funDy2);
        var result = dependsValueSplit.ElementAtOrDefault(int.Parse(funDy1) - 1) ?? string.Empty;
        return result;
    }

    public static string GetDic(
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "物品编号" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "2" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? funDy1 : funDy3;

        if (!exportWildcardDyData.TryGetValue(funDy1, out var val)) return string.Empty;
        var baseDicKey = val.Substring(0, val.Length - int.Parse(funDy2)) + funDy3;
        if (!strDictionary.TryGetValue(funDepends, out var dependsDicValue)) return string.Empty;
        if (!dependsDicValue.TryGetValue(baseDicKey, out var dependsValueList)) return string.Empty;

        var baseNum = val;

        if (dependsValueList.Contains(baseNum))
        {
            return string.Join(",", dependsValueList);
        }
        return string.Empty;
    }

    public static string GetDicKey(string funDepends)
    {
        string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        string filePath = Path.Combine(documentsPath, "strDic.csv");
        var fileDicData = LoadDictionaryFromFile(filePath);
        if (!fileDicData.TryGetValue(funDepends, out var dependsDicValue)) return string.Empty;
        return string.Join(",", dependsDicValue.Keys);
    }

    public static string SplitArr(Dictionary<string, string> exportWildcardDyData, string funDepends, string funDy1, string funDy2)
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "#" : funDy2;
        var dependsValue = exportWildcardDyData.GetValueOrDefault(funDepends, "");
        var dependsValueSplit = Regex.Split(dependsValue, funDy2);
        var result = dependsValueSplit.ElementAtOrDefault(int.Parse(funDy1) - 1) ?? string.Empty;
        return result;
    }

    public static string CollectRow(
        Dictionary<string, string> exportWildcardDyData,
        string funDepends,
        string funDy1,
        string funDy2,
        string funDy3,
        string funDy4,
        string funDy5,
        Dictionary<string, List<string>> baseData,
        string id
    )
    {
        funDy1 = string.IsNullOrEmpty(funDy1) ? "1" : funDy1;
        funDy2 = string.IsNullOrEmpty(funDy2) ? "消耗ID组" : funDy2;
        funDy3 = string.IsNullOrEmpty(funDy3) ? "消耗量组" : funDy3;
        funDy4 = string.IsNullOrEmpty(funDy4) ? "20" : funDy4;
        funDy5 = string.IsNullOrEmpty(funDy5) ? "1" : funDy5;

        List<string> idList = baseData.ContainsKey(id) ? baseData[id] : new List<string>();
        List<string> funDy2List = baseData.ContainsKey(funDy2) ? baseData[funDy2] : new List<string>();
        List<string> funDy3List = baseData.ContainsKey(funDy3) ? baseData[funDy3] : new List<string>();

        var loopTimes = int.Parse(funDy4);
        if (!exportWildcardDyData.TryGetValue(funDepends, out var dependsVal)) return string.Empty;
        if (!long.TryParse(dependsVal, out long collectRowId))
        {
            Debug.Print($"CollectRow: 无法将 '{dependsVal}' 解析为 long 类型。");
            return dependsVal;
        }

        string strCollect = string.Empty;
        string spawnCollect = string.Empty;

        // 首次的数据
        var idCollect = collectRowId.ToString();
        int findIndexFirst = idList.FindIndex(f => f == collectRowId.ToString());
        if (findIndexFirst != -1)
        {
            var funDy2Str = funDy2List[findIndexFirst];
            var funDy3Str = funDy3List[findIndexFirst];
            if (!string.IsNullOrEmpty(funDy2Str))
            {
                var funDy2StrSplit = Regex.Split(funDy2Str, "#");
                var funDy3StrSplit = Regex.Split(funDy3Str, "#");
                if (funDy3StrSplit.Length == funDy2StrSplit.Length)
                {
                    var sb = new StringBuilder();
                    for (int j = 0; j < funDy3StrSplit.Length; j++)
                    {
                        var temp = $"[{funDy2StrSplit[j]},{funDy3StrSplit[j]},{funDy2StrSplit[j]}]";
                        sb.Append(temp).Append(',');
                    }
                    if (sb.Length > 0) sb.Length--;
                    strCollect = sb.ToString();
                }
            }
        }
        if (string.IsNullOrEmpty(strCollect))
        {
            Debug.Print($"{idCollect}消耗数据为空，无法导出");
            return string.Empty;
        }

        strCollect = $"[{strCollect}]";

        // 其他次数据
        for (int i = 0; i < loopTimes; i++)
        {
            string stringSubCollect = string.Empty;
            collectRowId += int.Parse(funDy1);
            int findIndex = idList.FindIndex(f => f == collectRowId.ToString());
            if (findIndex != -1)
            {
                var funDy2Str = funDy2List[findIndex];
                var funDy3Str = funDy3List[findIndex];
                if (!string.IsNullOrEmpty(funDy2Str))
                {
                    var funDy2StrSplit = Regex.Split(funDy2Str, "#");
                    var funDy3StrSplit = Regex.Split(funDy3Str, "#");
                    if (funDy3StrSplit.Length == funDy2StrSplit.Length)
                    {
                        var sb = new StringBuilder();
                        for (int j = 0; j < funDy3StrSplit.Length; j++)
                        {
                            var temp = $"[{funDy2StrSplit[j]},{funDy3StrSplit[j]},{funDy2StrSplit[j]}]";
                            sb.Append(temp).Append(',');
                        }
                        if (sb.Length > 0) sb.Length--;
                        stringSubCollect = sb.ToString();
                        strCollect += $",[{stringSubCollect}]";
                        idCollect += "," + collectRowId;
                    }
                    else
                    {
                        spawnCollect = collectRowId.ToString();
                        break;
                    }
                }
                else
                {
                    spawnCollect = collectRowId.ToString();
                    break;
                }
            }
            else
            {
                spawnCollect = collectRowId.ToString();
                break;
            }
        }

        if (funDy5 == "1") return $"[{idCollect}]";
        if (funDy5 == "2") return $"[{strCollect}]";
        if (funDy5 == "3") return spawnCollect;

        return exportWildcardDyData.GetValueOrDefault(funDepends, string.Empty);
    }

    public static void GetDyWildcardValue(
        Dictionary<string, List<string>> baseData,
        Dictionary<string, string> exportWildcardDyData,
        string wildcard,
        string funDepends,
        int idCount
    )
    {
        var wildcardValuePattern = "#";
        if (funDepends.Contains("Var"))
        {
            var wildcardValueSplit = Regex.Split(funDepends, wildcardValuePattern);
            string fixWildcardValue = baseData.GetValueOrDefault(wildcardValueSplit[1], new List<string>()).ElementAtOrDefault(idCount) ?? string.Empty;
            if (wildcardValueSplit.Length == 3)
            {
                fixWildcardValue = fixWildcardValue.Replace(wildcardValuePattern, wildcardValueSplit[2]);
            }
            exportWildcardDyData[wildcard] = fixWildcardValue;
        }
    }

    public static void InitializeDictionary(
        Dictionary<string, Dictionary<string, List<string>>> strDictionary,
        string key,
        string subKey
    )
    {
        if (!strDictionary.ContainsKey(key))
        {
            strDictionary[key] = new Dictionary<string, List<string>>();
        }
        if (!strDictionary[key].ContainsKey(subKey))
        {
            strDictionary[key][subKey] = new List<string>();
        }
    }

    public static List<int> LoopNumber(int start, int max)
    {
        var sequence = new List<int>();
        for (int i = 1; i <= max; i++)
        {
            var modValue = ((start - 1) % max) + 1;
            start++;
            sequence.Add(modValue);
        }
        return sequence;
    }

    public static void SaveDictionaryToFile(
        Dictionary<string, Dictionary<string, List<string>>> dictionary,
        string filePath
    )
    {
        using StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8);
        foreach (var outerPair in dictionary)
        {
            foreach (var innerPair in outerPair.Value)
            {
                var line = $"{outerPair.Key},{innerPair.Key},{string.Join(",", innerPair.Value)}";
                writer.WriteLine(line);
            }
        }
    }

    public static Dictionary<string, Dictionary<string, List<string>>> LoadDictionaryFromFile(string filePath)
    {
        var dictionary = new Dictionary<string, Dictionary<string, List<string>>>();
        if (!File.Exists(filePath)) return dictionary;

        using StreamReader reader = new StreamReader(filePath, Encoding.UTF8);
        string line;
        while ((line = reader.ReadLine()) != null)
        {
            var parts = line.Split(',');
            if (parts.Length < 2) continue;
            string outerKey = parts[0];
            string innerKey = parts[1];
            List<string> values = parts.Length > 2 ? new List<string>(parts[2..]) : new List<string>();
            if (!dictionary.ContainsKey(outerKey)) dictionary[outerKey] = new Dictionary<string, List<string>>();
            dictionary[outerKey][innerKey] = values;
        }

        return dictionary;
    }
}
