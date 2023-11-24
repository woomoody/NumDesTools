using System;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using NPOI.SS.Formula.Functions;

namespace NumDesTools.image;

public class TmCaculate
{
    private static readonly dynamic Wk = CreatRibbon._app.ActiveWorkbook;
    private static readonly dynamic Path = Wk.Path;
    //TM关卡目标生成
    public static void CreatTmTargetEle()
    {
        var ws = Wk.ActiveSheet;
        Range modelRange = ws.Range["L4:Q2003"];
        var wsEle = Wk.Worksheets["TM元素"];
        Range targetEleMax = wsEle.Range["N16:R25"];
        var wsNewEle = Wk.Worksheets["TM关卡设计"];
        // 读取数据到一个二维数组中
        object[,] modelRangeValue = modelRange.Value ;
        var modelRangeValueList = PubMetToExcel.RangeDataToList(modelRangeValue);
        object[,] targetEleMaxValue = targetEleMax.Value;
        var targetEleMaxValueList = PubMetToExcel.RangeDataToList(targetEleMaxValue);
        //计数列表
        Dictionary<string, int> eleCount = new Dictionary<string, int>();
        //新eleID列表
        List<List<int>> targetRangeValueList = new List<List<int>>();
        for (var i = 0; i < modelRangeValueList.Count; i++)
        {
            var tempTarget = new List<int>();
            for (var j = 0; j < modelRangeValueList[i].Count; j++)
            {
                if (modelRangeValueList[i][j] != null)
                {
                    string ele = modelRangeValueList[i][j].ToString();
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
        Range targetEleMax = wsEle.Range["N16:R25"];
        var wsNewEle = Wk.Worksheets["TM关卡设计"];
        Range targetModelRange = wsNewEle.Range["B2:G2001"];
        // 读取数据到一个二维数组中
        object[,] modelRangeValue = modelRange.Value;
        var modelRangeValueList = PubMetToExcel.RangeDataToList(modelRangeValue);
        object[,] modelRangeValue2 = modelRange2.Value;
        var modelRangeValueList2 = PubMetToExcel.RangeDataToList(modelRangeValue2);
        object[,] targetEleMaxValue = targetEleMax.Value;
        var targetEleMaxValueList = PubMetToExcel.RangeDataToList(targetEleMaxValue);
        object[,] targetModelRangeValue = targetModelRange.Value;
        var targetModelRangeValueList = PubMetToExcel.RangeDataToList(targetModelRangeValue);
        //计数列表
        Dictionary<string, int> eleCount = new Dictionary<string, int>();
        //新eleID列表
        List<List<int>> targetRangeValueList = new List<List<int>>();
        for (var i = 0; i < modelRangeValueList.Count; i++)
        {
            var tempTarget = new List<int>();
            for (var j = 0; j < modelRangeValueList[i].Count; j++)
            {
                if (modelRangeValueList[i][j] != null)
                {
                    string ele = modelRangeValueList[i][j].ToString();
                    if (eleCount.ContainsKey(ele))
                    {
                        eleCount[ele]++;
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
                            //Id去目标重
                            for (var n = 0; n < modelRangeValueList2[i].Count; n++)
                            {
                                string targetEle = modelRangeValueList2[i][n].ToString();
                                if (targetEle == ele)
                                {
                                    eleId = Convert.ToInt32(targetModelRangeValueList[i][n]);
                                }
                            }
                            var eleMax = Convert.ToInt32(targetEleMaxValueList[k][4]);
                            var targetId = eleId + (eleCount[ele] - 1) % eleMax + 1;
                            tempTarget.Add(targetId);
                        }
                    }
                }
            }
            targetRangeValueList.Add(tempTarget);
        }
        //写入数据
        PubMetToExcel.ListToArrayToRange(targetRangeValueList, wsNewEle, 2, 9);
    }
}