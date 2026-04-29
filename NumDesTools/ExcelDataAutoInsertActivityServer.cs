using System.Text;
using MiniExcelLibs;
using MessageBox = System.Windows.MessageBox;

#pragma warning disable CA1416

namespace NumDesTools;

public static class ExcelDataAutoInsertActivityServer
{
    private const double SecondsInADay = 86400;
    private const double OneMinuteInDays = 60 / SecondsInADay;

    public static void Source(bool isNames)
    {
        var indexWk = NumDesAddIn.App.ActiveWorkbook;

        var sourceSheet = indexWk.Worksheets["运营排期"];
        var targetSheet = indexWk.Worksheets["Sheet1"];
        var fixSheet = indexWk.Worksheets["活动模板"];
        var lifeTypeSheet = indexWk.Worksheets["生命周期"];

        var fixData = PubMetToExcel.ExcelDataToList(fixSheet);
        var fixTitle = fixData.Item1;
        List<List<object>> fixDataList = fixData.Item2;
        //删除活动名或者活动id列为空的数据
        fixDataList = fixDataList.Where(row => row[0] != null && row[1] != null).ToList();

        var fixNames = fixTitle.IndexOf("活动名称");
        var fixIds = fixTitle.IndexOf("活动id");
        var fixPush = fixTitle.IndexOf("前端可获取活动时间");
        //var fixPushEnds = fixTitle.IndexOf("停止向前端发送活动时间");
        //var fixPreHeats = fixTitle.IndexOf("预热期开始时间");
        //var fixOpens = fixTitle.IndexOf("活动开启时间");
        //var fixEnds = fixTitle.IndexOf("活动结束时间");
        var fixCloses = fixTitle.IndexOf("活动关闭时间");
        var isActGroup = fixTitle.IndexOf("是否活动组");
        var openCondition = fixTitle.IndexOf("活动开启条件");
        var lifeType = fixTitle.IndexOf("生命周期类型");

        var lifeTypeData = PubMetToExcel.ExcelDataToList(lifeTypeSheet);
        var lifeTypeTitle = lifeTypeData.Item1;
        List<List<object>> lifeTypeDataList = lifeTypeData.Item2;
        var lifeTypeIndex = lifeTypeTitle.IndexOf("类型");
        var lifeTypeValue = lifeTypeTitle.IndexOf("内容");

        var sourceMaxCol = sourceSheet.UsedRange.Columns.Count;
        var sourceMaxRow = sourceSheet.UsedRange.Rows.Count;
        var sourceRange = sourceSheet.Range[
            sourceSheet.Cells[3, 5],
            sourceSheet.Cells[sourceMaxRow, sourceMaxCol]
        ];
        var sourceDateRange = sourceSheet.Range[
            sourceSheet.Cells[3, 3],
            sourceSheet.Cells[sourceMaxRow, 3]
        ];
        var sourceOutRange = sourceSheet.Range[
            sourceSheet.Cells[2, 5],
            sourceSheet.Cells[2, sourceMaxCol]
        ];

        int nameOrId = isNames ? fixNames : fixIds;
        string nameOrIdString = isNames ? "活动名" : "活动ID";

        Array sourceDataArr = sourceDateRange.Value2;
        var sourceData = new List<(string, double, double, int, int, int, string)>();

        for (int col = 1; col <= sourceMaxCol - 3 + 1; col++)
        {
            for (int row = 1; row <= sourceMaxRow - 3 + 1; row++)
            {
                var cell = sourceRange[row, col];

                // 过滤已删除活动
                bool hasStrikethrough = cell.Font.Strikethrough;
                if (hasStrikethrough)
                    continue;

                var cellOutValue = sourceOutRange[1, col].Value2?.ToString() ?? "";
                if (cellOutValue != "#导出")
                    continue;

                if (cell.MergeCells)
                {
                    var mergeRange = cell.MergeArea;
                    if (cell.Address == mergeRange.Cells[1, 1].Address)
                    {
                        var mergeValue = mergeRange.Cells[1, 1].Value2;
                        if (mergeValue == null)
                            continue;
                        var activityName = mergeValue.ToString();
                        var activityCondition = "";
                        if (activityName.Contains("："))
                        {
                            var parts = activityName.Split("：");
                            activityCondition = parts[0];
                            activityName = parts.Length > 1 ? parts[1] : string.Empty;
                        }

                        sourceData.Add(
                            (
                                activityName,
                                (double)sourceDataArr.GetValue(mergeRange.Row - 2, 1),
                                (double)
                                    sourceDataArr.GetValue(
                                        mergeRange.Row + mergeRange.Rows.Count - 3,
                                        1
                                    ),
                                mergeRange.Column,
                                mergeRange.Row,
                                mergeRange.Row + mergeRange.Rows.Count - 1,
                                activityCondition
                            )
                        );
                    }
                }
                else if (cell.Value != null)
                {
                    var activityName = cell.Value.ToString();
                    var activityCondition = "";
                    if (activityName.Contains("："))
                    {
                        var parts = activityName.Split("：");
                        activityCondition = parts[0];
                        activityName = parts.Length > 1 ? parts[1] : string.Empty;
                    }
                    sourceData.Add(
                        (
                            activityName,
                            (double)sourceDataArr.GetValue(cell.Row - 2, 1),
                            (double)sourceDataArr.GetValue(cell.Row + cell.Rows.Count - 3, 1),
                            cell.Column,
                            cell.Row,
                            cell.Row + cell.Rows.Count - 1,
                            activityCondition
                        )
                    );
                }
            }
        }

        var targetDataList = new List<List<string>>();
        var errorLog = new StringBuilder();
        var unixEpoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        foreach (var a in sourceData)
        {
            var fixDataMatch = fixDataList.FirstOrDefault(b => b[nameOrId].ToString() == a.Item1);
            if (fixDataMatch == null)
            {
                var activeName = a.Item1;
                if (a.Item7 != "")
                {
                    activeName = $"{a.Item7}：{a.Item1}";
                }
                errorLog.Append($"运营排期-未找到-活动模版【{nameOrIdString}】：{activeName}\r\n");
                targetDataList.Add(
                    [
                        "targetId",
                        a.Item1,
                        "targetPushTimeString",
                        "targetPushTimeLong",
                        "targetPushEndTimeString",
                        "targetPushEndTimeLong",
                        "targetPreHeatTimeString",
                        "targetPreHeatTimeLong",
                        "targetOpenTimeString",
                        "targetOpenTimeLong",
                        "targetEndTimeString",
                        "targetEndTimeLong",
                        "targetCloseTimeString",
                        "targetCloseTimeLong",
                        "targetActGroup",
                        "targetOpenCondition",
                        "targetLifeType"
                    ]
                );
                continue;
            }

            var sourceStartTimeLong = (long)
                (DateTime.FromOADate(a.Item2).ToUniversalTime() - unixEpoch).TotalSeconds;
            var sourceEndTimeLong = (long)
                (
                    DateTime
                        .FromOADate(a.Item2 + a.Item6 - a.Item5 + 1 - OneMinuteInDays)
                        .ToUniversalTime() - unixEpoch
                ).TotalSeconds;

            string ConvertToDateString(double oaDate, object hoursOffset)
            {
                double hoursOffsetDou = Convert.ToDouble(hoursOffset);
                return DateTime
                    .FromOADate(oaDate)
                    .AddHours(hoursOffsetDou * 24)
                    .ToString(CultureInfo.InvariantCulture);
            }

            long ConvertToUnixTime(long baseTime, object hoursOffset)
            {
                double hoursOffsetDou = Convert.ToDouble(hoursOffset);
                return baseTime + (long)(hoursOffsetDou * 24 * 3600);
            }

            var targetId = fixDataMatch[fixIds].ToString();
            var targetName = a.Item1;
            if (a.Item7 != "")
            {
                targetName = $"{a.Item7}：{a.Item1}";
            }
            var targetPushTimeString = ConvertToDateString(a.Item2, fixDataMatch[fixPush]);
            var targetPushTimeLong = ConvertToUnixTime(sourceStartTimeLong, fixDataMatch[fixPush]);
            var targetPushEndTimeString = ConvertToDateString(
                a.Item2 + a.Item6 - a.Item5 + 1 - OneMinuteInDays,
                0
            //fixDataMatch[fixPushEnds]
            );
            var targetPushEndTimeLong = ConvertToUnixTime(
                sourceEndTimeLong,
                0
            //fixDataMatch[fixPushEnds]
            );
            var targetPreHeatTimeString = ConvertToDateString(
                a.Item2,
                0 /*fixDataMatch[fixPreHeats]*/
            );
            var targetPreHeatTimeLong = ConvertToUnixTime(
                sourceStartTimeLong,
                0
            //fixDataMatch[fixPreHeats]
            );
            var targetOpenTimeString = ConvertToDateString(
                a.Item2,
                0 /*fixDataMatch[fixOpens]*/
            );
            var targetOpenTimeLong = ConvertToUnixTime(
                sourceStartTimeLong,
                0 /*fixDataMatch[fixOpens]*/
            );
            var targetEndTimeString = ConvertToDateString(
                a.Item3 - OneMinuteInDays,
                0
                    /*fixDataMatch[fixEnds] */+ 1
            );
            var targetEndTimeLong = ConvertToUnixTime(
                sourceEndTimeLong,
                0 /*fixDataMatch[fixEnds]*/
            );
            var targetCloseTimeString = ConvertToDateString(
                a.Item3 - OneMinuteInDays,
                fixDataMatch[fixCloses] + 1
            );
            var targetCloseTimeLong = ConvertToUnixTime(sourceEndTimeLong, fixDataMatch[fixCloses]);
            var targetActGroup = fixDataMatch[isActGroup].ToString();
            var targetOpenCondition = fixDataMatch[openCondition]?.ToString() ?? "";
            if (targetOpenCondition == "\"{}\"" || targetOpenCondition == "")
            {
                if (a.Item7 != "")
                {
                    targetOpenCondition = "\"{{26,{" + a.Item7.Replace("、", ",") + "}}}\"";
                }
            }
            var targetLifeType = fixDataMatch[lifeType];
            string targetLifeValue;
            if (targetLifeType == null)
            {
                targetLifeValue = "";
            }
            else
            {
                var lifeTypeMatch = lifeTypeDataList.FirstOrDefault(l =>
                    l[lifeTypeIndex].ToString() == targetLifeType.ToString()
                );
                if (lifeTypeMatch == null)
                {
                    var activeName = a.Item1;
                    if (a.Item7 != "")
                    {
                        activeName = $"{a.Item7}：{a.Item1}";
                    }
                    errorLog.Append($"运营排期-活动模版【{nameOrIdString}】：{activeName}**生命周期类型错误[{targetLifeType}]，搜索不到\r\n");
                    targetLifeValue = "targetLifeValue";
                }
                else
                {
                    targetLifeValue = lifeTypeMatch[lifeTypeValue]?.ToString() ?? "";
                }
            }
            targetDataList.Add(
                [
                    targetId,
                    targetName,
                    targetPushTimeString,
                    targetPushTimeLong.ToString(),
                    targetPushEndTimeString,
                    targetPushEndTimeLong.ToString(CultureInfo.InvariantCulture),
                    targetPreHeatTimeString,
                    targetPreHeatTimeLong.ToString(),
                    targetOpenTimeString,
                    targetOpenTimeLong.ToString(),
                    targetEndTimeString,
                    targetEndTimeLong.ToString(CultureInfo.InvariantCulture),
                    targetCloseTimeString,
                    targetCloseTimeLong.ToString(CultureInfo.InvariantCulture),
                    targetActGroup,
                    targetOpenCondition,
                    targetLifeValue
                ]
            );
        }

        if (errorLog.Length > 0)
        {
            ErrorLogCtp.DisposeCtp();
            ErrorLogCtp.CreateCtp(errorLog.ToString());
            MessageBox.Show(@"有活动找不到，查看错误日志");
            sourceSheet.Select();
        }
        else
        {
            targetSheet.Select();
        }
        var targetStartCol = 2;
        var targetStartRow = 5;
        var targetRangeOld = targetSheet.Range[
            targetSheet.Cells[targetStartRow, targetStartCol],
            targetSheet.Cells[targetSheet.UsedRange.Rows.Count, targetSheet.UsedRange.Columns.Count]
        ];
        targetRangeOld.Value = null;

        var rows = targetDataList.Count;
        var columns = targetDataList[0].Count;
        var targetDataArr = new string[rows, columns];
        for (var i = 0; i < rows; i++)
        {
            for (var j = 0; j < columns; j++)
            {
                targetDataArr[i, j] = targetDataList[i][j];
            }
        }

        var targetRange = targetSheet.Range[
            targetSheet.Cells[targetStartRow, targetStartCol],
            targetSheet.Cells[
                targetStartRow + targetDataArr.GetLength(0) - 1,
                targetStartCol + targetDataArr.GetLength(1) - 1
            ]
        ];
        targetRange.Value = targetDataArr;
    }

    public static void ModeDataUpdate()
    {
        var wk = NumDesAddIn.App.ActiveWorkbook;
        var basePath = wk.Path;

        var baseList = PubMetToExcel.GetExcelListObjects("活动枚举", "活动枚举");
        if (baseList == null)
        {
            MessageBox.Show("活动枚举 中的名称表-【活动枚举】不存在");
            return;
        }
        if (baseList.DataBodyRange == null) { MessageBox.Show("活动枚举表无数据行"); return; }
        object[,] baseArray = baseList.DataBodyRange.Value2;
        var baseDic = PubMetToExcel.TwoDArrayToDictionaryFirstKey1(baseArray);

        var activityGroup = basePath + @"\ActivityClientHierarchyGroupData.xlsx";
        var activityGroupSheetName = "Sheet1";
        var activityGroupData = MiniExcel.Query(
            activityGroup,
            sheetName: activityGroupSheetName,
            startCell: "B2",
            useHeaderRow: true
        );
        var activityGroupSub = basePath + @"\ActivityClientHierarchyData.xlsx";
        var activityGroupSubSheetName = "Sheet1";
        var activityGroupSubData = MiniExcel.Query(
            activityGroupSub,
            sheetName: activityGroupSubSheetName,
            startCell: "B2",
            useHeaderRow: true
        );
        var activty = basePath + @"\ActivityClientData.xlsx";
        var activtySheetName = "Sheet1";
        var activtyData = MiniExcel.Query(
            activty,
            sheetName: activtySheetName,
            startCell: "B2",
            useHeaderRow: true
        );

        // 预先处理activityData，建立id到type的映射
        var activityDataMap = new Dictionary<string, List<string>>();
        var activityIdList = new List<string>();
        foreach (var activtyRow in activtyData.Skip(3))
        {
            if (activtyRow is IDictionary<string, object> activtyRowDict)
            {
                string id = activtyRowDict["id"]?.ToString();
                string type = activtyRowDict["type"]?.ToString();
                string comment = activtyRowDict["#备注"]?.ToString();

                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(type))
                {
                    activityDataMap[id] = new List<string> { comment, type };
                }

                activityIdList.Add(id);
            }
        }

        // 预先处理activityGroupSubData，建立id到activityID的映射
        var subDataMap = new Dictionary<string, string>();
        foreach (var subRow in activityGroupSubData.Skip(3))
        {
            if (subRow is IDictionary<string, object> subRowDict)
            {
                string id = subRowDict["id"]?.ToString();
                string activityId = subRowDict["activityIds"]?.ToString();

                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(activityId))
                {
                    subDataMap[id] = activityId; // 假设id是唯一的，否则使用Add方法
                }
            }
        }

        var activityInfo = new Dictionary<string, List<string>>();
        var activityGroupIdList = new List<string>();
        // 处理activityGroupData
        foreach (var row in activityGroupData.Skip(3)) // 跳过前3行标题
        {
            if (row is IDictionary<string, object> rowDict)
            {
                string activityGroupId = rowDict["id"]?.ToString();
                string hierarchyActivityIDs = rowDict["hierarchyActivityIDs"]?.ToString();
                string activityGroupComment = rowDict["#备注"]?.ToString();

                if (
                    !string.IsNullOrEmpty(activityGroupId)
                    && !string.IsNullOrEmpty(hierarchyActivityIDs)
                )
                {
                    activityGroupIdList.Add(activityGroupId);

                    // 处理hierarchyActivityIDs格式：[id1,id2,id3]
                    var hierarchyActivityIDsNums = hierarchyActivityIDs
                        .Trim('[', ']')
                        .Split(',')
                        .Select(s => s.Trim())
                        .Where(s => !string.IsNullOrEmpty(s))
                        .ToList();

                    if (hierarchyActivityIDsNums.Count == 0)
                        continue;

                    foreach (var hierarchyActivityIDsNum in hierarchyActivityIDsNums)
                    {
                        // 通过活动id查找具体信息

                        string activityId = String.Empty;
                        if (subDataMap.ContainsKey(hierarchyActivityIDsNum))
                        {
                            activityId = subDataMap[hierarchyActivityIDsNum];
                        }

                        string activityComment = String.Empty;
                        string activityType = String.Empty;

                        if (activityDataMap.ContainsKey(activityId))
                        {
                            activityComment = activityDataMap[activityId][0];
                            activityType = activityDataMap[activityId][1];
                        }

                        List<string> activityBaseInfo;
                        if (baseDic.ContainsKey(activityType))
                        {
                            activityBaseInfo = baseDic[activityType].Skip(1).ToList();
                        }
                        else
                        {
                            activityBaseInfo = baseDic["通用"].Skip(1).ToList();
                        }
                        var activityAllInfo = new List<string>();
                        activityAllInfo.Add(activityId);
                        activityAllInfo.Add(activityComment);
                        activityAllInfo.AddRange(activityBaseInfo);
                        activityAllInfo.Add(activityType);
                        activityAllInfo.Add(activityGroupId); // 标记归属的活动组
                        if (activityAllInfo.Any())
                        {
                            activityInfo[activityId] = activityAllInfo;
                        }
                    }
                    if (activityInfo.ContainsKey(hierarchyActivityIDsNums[0]))
                    {
                        var activityGroupAllInfo = new List<string>(
                            activityInfo[hierarchyActivityIDsNums[0]]
                        );
                        activityGroupAllInfo[3] = "1"; // 设置为活动组
                        activityGroupAllInfo[0] = activityGroupId;
                        activityGroupAllInfo[1] = activityGroupComment;
                        if (activityGroupAllInfo.Any())
                        {
                            activityInfo[activityGroupId] = activityGroupAllInfo;
                        }
                    }
                }
            }
        }

        // 检查重复ID
        var repeatKeyList = activityIdList.Intersect(activityGroupIdList).ToList();
        if (repeatKeyList.Count > 0)
        {
            string repeatKeys = string.Join(", ", repeatKeyList);
            MessageBox.Show($"存在重复活动ID，无法继续写入：{repeatKeys}");
            return;
        }

        // 处理剩余activityData
        foreach (var activity in activityDataMap)
        {
            string activityId = activity.Key;
            string activityComment;
            string activityType;

            if (!activityInfo.ContainsKey(activityId))
            {
                activityComment = activityDataMap[activityId][0];
                activityType = activityDataMap[activityId][1];
                List<string> activityBaseInfo;
                if (baseDic.ContainsKey(activityType))
                {
                    activityBaseInfo = baseDic[activityType].Skip(1).ToList();
                }
                else
                {
                    activityBaseInfo = baseDic["通用"].Skip(1).ToList();
                }
                var activityAllInfo = new List<string>();
                activityAllInfo.Add(activityId);
                activityAllInfo.Add(activityComment);
                activityAllInfo.AddRange(activityBaseInfo);
                activityAllInfo.Add(activityType);
                activityAllInfo.Add("无活动组"); // 标记归属的活动组
                if (activityAllInfo.Any())
                {
                    activityInfo[activityId] = activityAllInfo;
                }
            }
        }
        // 写入数据
        var activityArray = PubMetToExcel.DictionaryTo2DArray(activityInfo);

        var activityList = PubMetToExcel.GetExcelListObjects("活动模板", "活动模板");
        if (activityList == null)
        {
            MessageBox.Show("活动模板 中的名称表-【活动模板】不存在");
            return;
        }

        var rowMax = activityArray.GetLength(0);
        PubMetToExcel.WriteExcelDataC("活动模板", 1, 10000, 0, 8, null);
        PubMetToExcel.WriteExcelDataC("活动模板", 1, rowMax, 0, 8, activityArray);

        activityList.Resize(
            activityList.Range.Resize[rowMax + 1, activityList.Range.Columns.Count]
        );
    }
}
