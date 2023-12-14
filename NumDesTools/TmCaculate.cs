using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace NumDesTools;

public class TmCaculate
{
    private static readonly dynamic Wk = CreatRibbon.App.ActiveWorkbook;
    //TM关卡目标生成
    public static void CreatTmTargetEle()
    {
        var ws = Wk.ActiveSheet;
        Range modelRange = ws.Range["L4:Q2003"];
        var wsEle = Wk.Worksheets["TM元素"];
        Range targetEleMax = wsEle.Range["N16:S25"];
        var wsNewEle = Wk.Worksheets["TM关卡设计"];
        // 读取数据到一个二维数组中
        object[,] modelRangeValue = modelRange.Value ;
        var modelRangeValueList = PubMetToExcel.RangeDataToList(modelRangeValue);
        object[,] targetEleMaxValue = targetEleMax.Value;
        var targetEleMaxValueList = PubMetToExcel.RangeDataToList(targetEleMaxValue);
        //计数列表
        var eleCount = new Dictionary<string, int>();
        //新eleID列表
        var targetRangeValueList = new List<List<object>>();
        for (var i = 0; i < modelRangeValueList.Count; i++)
        {
            var tempTarget = new List<object>();
            for (var j = 0; j < modelRangeValueList[i].Count; j++)
            {
                if (modelRangeValueList[i][j] != null)
                {
                    var ele = modelRangeValueList[i][j].ToString();
                    if (eleCount.ContainsKey(ele))
                    {
                        eleCount[ele] ++;
                    }
                    else
                    {
                        eleCount[ele] = 1;
                    }
                    for (var k = 0; k < targetEleMaxValueList.Count; k++)
                    {
                        if (ele == targetEleMaxValueList[k][0].ToString())
                        {
                            var eleId = Convert.ToInt32(targetEleMaxValueList[k][3]);
                            var eleMax = Convert.ToInt32(targetEleMaxValueList[k][1]);
                            var targetId = eleId + (eleCount[ele] - 1) % eleMax +1;
                            tempTarget.Add(targetId);
                        }
                    }
                }
            }
            targetRangeValueList.Add(tempTarget);
        }
        //写入数据
        PubMetToExcel.ListToArrayToRange(targetRangeValueList, wsNewEle,2,2);
    }
    //TM关卡非目标生成
    public static void CreatTmNormalEle()
    {
        var ws = Wk.ActiveSheet;
        Range modelRange = ws.Range["R4:AP2003"];
        Range modelRange2 = ws.Range["L4:Q2003"];
        var wsEle = Wk.Worksheets["TM元素"];
        Range targetEleMax = wsEle.Range["N16:S25"];
        var wsNewEle = Wk.Worksheets["TM关卡设计"];
        Range targetModelRange = wsNewEle.Range["B2:G2001"];
        // 读取数据到一个二维数组中
        object[,] modelRangeValue = modelRange.Value;
        var modelRangeValueList = PubMetToExcel.RangeDataToList(modelRangeValue);
        object[,] modelRangeValue2 = modelRange2.Value;
        PubMetToExcel.RangeDataToList(modelRangeValue2);
        object[,] targetEleMaxValue = targetEleMax.Value;
        var targetEleMaxValueList = PubMetToExcel.RangeDataToList(targetEleMaxValue);
        object[,] targetModelRangeValue = targetModelRange.Value;
        var targetModelRangeValueList = PubMetToExcel.RangeDataToList(targetModelRangeValue);

        //新eleID列表
        var targetRangeValueList = new List<List<object>>();
        //非目标计数列表
        var eleCount = new Dictionary<string, int>();
        //非目标ID库循环列表
        var eleIdLoop = new Dictionary<string, List<int>>();
        for (var i = 0; i < targetEleMaxValueList.Count; i ++)
        {
            var loopTimes = Convert.ToInt32(targetEleMaxValueList[i][5]);
            var eleMax = Convert.ToInt32(targetEleMaxValueList[i][4]);
            var eleBaseId = Convert.ToInt32(targetEleMaxValueList[i][3]);
            var eleTheme = Convert.ToString(targetEleMaxValueList[i][0]);
            var loopIdList = new List<int>();
            for (var j = 0; j < loopTimes*eleMax; j++)
            {
                var loopId = (j+1) % eleMax;
                if (loopId == 0) loopId = eleMax;
                loopIdList.Add( eleBaseId + loopId);
            }
            eleIdLoop[eleTheme] = loopIdList;
        }
        for (var i = 0; i < modelRangeValueList.Count; i++)
        {
            var tempTarget = new List<object>();
            for (var j = 0; j < modelRangeValueList[i].Count; j++)
            {
                if ((string)modelRangeValueList[i][j] != "")
                {
                    //非目标元素整体计数
                    var ele = modelRangeValueList[i][j].ToString();
                    if (eleCount.ContainsKey(ele))
                    {
                        eleCount[ele]++;
                    }
                    else
                    {
                        eleCount[ele] = 1;
                    }
                    //非目标元素ID获取、和目标元素去重（向后索引）处理
                    var eleId = eleIdLoop[ele][eleCount[ele] - 1];
                    foreach (var id in targetModelRangeValueList[i])
                    {
                        if (Convert.ToInt32(id) == eleId)
                        {
                            eleCount[ele]++;
                            eleId = eleIdLoop[ele][eleCount[ele] - 1];
                        }
                    }
                    tempTarget.Add(eleId);
                }
            }
            targetRangeValueList.Add(tempTarget);
        }

        //写入数据
        PubMetToExcel.ListToArrayToRange(targetRangeValueList, wsNewEle, 2, 9);
    }
}