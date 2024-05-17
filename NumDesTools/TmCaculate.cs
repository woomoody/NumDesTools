
namespace NumDesTools;

public class TmCaculate
{
    private static readonly dynamic Wk = NumDesAddIn.App.ActiveWorkbook;

    public static void CreatTmTargetEle()
    {
        var ws = Wk.ActiveSheet;
        Range modelRange = ws.Range["L4:Q2003"];
        var wsEle = Wk.Worksheets["TM元素"];
        Range targetEleMax = wsEle.Range["N16:S25"];
        var wsNewEle = Wk.Worksheets["TM关卡设计"];
        object[,] modelRangeValue = modelRange.Value;
        var modelRangeValueList = PubMetToExcel.RangeDataToList(modelRangeValue);
        object[,] targetEleMaxValue = targetEleMax.Value;
        var targetEleMaxValueList = PubMetToExcel.RangeDataToList(targetEleMaxValue);
        var eleCount = new Dictionary<string, int>();
        var targetRangeValueList = new List<List<object>>();
        foreach (var t in modelRangeValueList)
        {
            var tempTarget = new List<object>();
            for (var j = 0; j < t.Count; j++)
                if (t[j] != null)
                {
                    var ele = t[j].ToString();
                    if (ele != null && eleCount.ContainsKey(ele))
                        eleCount[ele]++;
                    else if (ele != null) eleCount[ele] = 1;
                    foreach (var t1 in targetEleMaxValueList)
                        if (ele == t1[0].ToString())
                        {
#pragma warning disable CA1305
                            var eleId = Convert.ToInt32(t1[3]);
#pragma warning restore CA1305
#pragma warning disable CA1305
                            var eleMax = Convert.ToInt32(t1[1]);
#pragma warning restore CA1305
                            if (ele != null)
                            {
                                var targetId = eleId + (eleCount[ele] - 1) % eleMax + 1;
                                tempTarget.Add(targetId);
                            }
                        }
                }

            targetRangeValueList.Add(tempTarget);
        }

        PubMetToExcel.ListToArrayToRange(targetRangeValueList, wsNewEle, 2, 2);
    }

    public static void CreatTmNormalEle()
    {
        var ws = Wk.ActiveSheet;
        Range modelRange = ws.Range["R4:AP2003"];
        Range modelRange2 = ws.Range["L4:Q2003"];
        var wsEle = Wk.Worksheets["TM元素"];
        Range targetEleMax = wsEle.Range["N16:S25"];
        var wsNewEle = Wk.Worksheets["TM关卡设计"];
        Range targetModelRange = wsNewEle.Range["B2:G2001"];
        object[,] modelRangeValue = modelRange.Value;
        var modelRangeValueList = PubMetToExcel.RangeDataToList(modelRangeValue);
        object[,] modelRangeValue2 = modelRange2.Value;
        PubMetToExcel.RangeDataToList(modelRangeValue2);
        object[,] targetEleMaxValue = targetEleMax.Value;
        var targetEleMaxValueList = PubMetToExcel.RangeDataToList(targetEleMaxValue);
        object[,] targetModelRangeValue = targetModelRange.Value;
        var targetModelRangeValueList = PubMetToExcel.RangeDataToList(targetModelRangeValue);

        var targetRangeValueList = new List<List<object>>();
        var eleCount = new Dictionary<string, int>();
        var eleIdLoop = new Dictionary<string, List<int>>();
        foreach (var t in targetEleMaxValueList)
        {
#pragma warning disable CA1305
            var loopTimes = Convert.ToInt32(t[5]);
#pragma warning restore CA1305
#pragma warning disable CA1305
            var eleMax = Convert.ToInt32(t[4]);
#pragma warning restore CA1305
#pragma warning disable CA1305
            var eleBaseId = Convert.ToInt32(t[3]);
#pragma warning restore CA1305
#pragma warning disable CA1305
            var eleTheme = Convert.ToString(t[0]);
#pragma warning restore CA1305
            var loopIdList = new List<int>();
            for (var j = 0; j < loopTimes * eleMax; j++)
            {
                var loopId = (j + 1) % eleMax;
                if (loopId == 0) loopId = eleMax;
                loopIdList.Add(eleBaseId + loopId);
            }

            if (eleTheme != null) eleIdLoop[eleTheme] = loopIdList;
        }

        for (var i = 0; i < modelRangeValueList.Count; i++)
        {
            var tempTarget = new List<object>();
            for (var j = 0; j < modelRangeValueList[i].Count; j++)
                if ((string)modelRangeValueList[i][j] != "")
                {
                    var ele = modelRangeValueList[i][j].ToString();
                    if (ele != null && eleCount.ContainsKey(ele))
                        eleCount[ele]++;
                    else if (ele != null) eleCount[ele] = 1;
                    if (ele != null)
                    {
                        var eleId = eleIdLoop[ele][eleCount[ele] - 1];
                        foreach (var id in targetModelRangeValueList[i])
#pragma warning disable CA1305
                            if (Convert.ToInt32(id) == eleId)
                            {
                                eleCount[ele]++;
                                eleId = eleIdLoop[ele][eleCount[ele] - 1];
                            }
#pragma warning restore CA1305

                        tempTarget.Add(eleId);
                    }
                }

            targetRangeValueList.Add(tempTarget);
        }

        PubMetToExcel.ListToArrayToRange(targetRangeValueList, wsNewEle, 2, 9);
    }
}