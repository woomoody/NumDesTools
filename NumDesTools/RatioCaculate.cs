using System.Text.RegularExpressions;

namespace NumDesTools;

/// <summary>
/// 概率计算类
/// </summary>
// ReSharper disable once UnusedMember.Global
public class RatioCaculate
{
    //卡牌收集概率计算
    // ReSharper disable once UnusedMember.Global
    public static void CacCardCollect()
    {
        //获取表格源数据
        var workBook = NumDesAddIn.App.ActiveWorkbook;
        var collectCardGroup = workBook.Worksheets["CollectCardGroup"];
        var collectCardInfo = workBook.Worksheets["CollectCardInfo"];
        var collectCardRarity = workBook.Worksheets["CollectCardRarity"];
        var collectCardScore = workBook.Worksheets["CollectCardScore"];

        var collectCardGroupData = PubMetToExcel.ExcelDataToListBySelf(collectCardGroup, 5, 1, 2, 1);
        var collectCardGroupTitle = collectCardGroupData.Item1;
        var collectCardGroupDataList = collectCardGroupData.Item2;
        var collectCardInfoData = PubMetToExcel.ExcelDataToListBySelf(collectCardInfo, 5, 1, 2, 1);
        var collectCardInfoTitle = collectCardInfoData.Item1;
        var collectCardInfoDataList = collectCardInfoData.Item2;
        var collectCardRarityData = PubMetToExcel.ExcelDataToListBySelf(collectCardRarity, 5, 1, 2, 1);
        var collectCardRarityTitle = collectCardRarityData.Item1;
        var collectCardRarityDataList = collectCardRarityData.Item2;
        var collectCardScoreData = PubMetToExcel.ExcelDataToListBySelf(collectCardScore, 5, 1, 2, 1);
        var collectCardScoreTitle = collectCardScoreData.Item1;
        var collectCardScoreDataLIst = collectCardScoreData.Item2;
        //分拆出每个卡组的稀有度构成
        var cardGroupIndex = collectCardGroupTitle.IndexOf("card_id");
        var cardGroupNameIndex = collectCardGroupTitle.IndexOf("#备注");
        var cardInfoIdIndex = collectCardInfoTitle.IndexOf("id");
        var cardInfoRarityIndex = collectCardInfoTitle.IndexOf("rarity");
        var cardRarityWeightIndex = collectCardRarityTitle.IndexOf("rate");
        var cardRarityScoreIndex = collectCardRarityTitle.IndexOf("reward");
        var cardScoreIndex = collectCardScoreTitle.IndexOf("parameter");

        var groupCount = collectCardGroupDataList.Count;
        var cardCount = collectCardInfoDataList.Count;
        var groupRarityCount = new List<(string, int, int, int)>();
        for (var i = 0; i < groupCount; i++)
        {
            var cardGroupStr = collectCardGroupDataList[i][cardGroupIndex] as string;
            if (cardGroupStr == null) continue; // 安全转换和处理
            var cardGroupName = collectCardGroupDataList[i][cardGroupNameIndex] as string;
            if (string.IsNullOrEmpty(cardGroupName)) continue;
            //拆ID，查ID，获取各个品质的个数
            var cardIdPattern = "\\d+";
            var cardIdMatches = Regex.Matches(cardGroupStr, cardIdPattern);
            if (cardIdMatches.Count == 0) continue;
            var rarity1 = 0;
            var rarity2 = 0;
            var rarity3 = 0;
            foreach (Match item in cardIdMatches)
            {
                var sourceCardId = item.Value;
                for (var j = 0; j < cardCount; j++)
                {
                    var targetCardId = collectCardInfoDataList[j][cardInfoIdIndex].ToString();
                    if (targetCardId == sourceCardId)
                    {
                        var targetCardRarity = (int)collectCardInfoDataList[j][cardInfoRarityIndex];
                        if (targetCardRarity == 1)
                            rarity1++;
                        else if (targetCardRarity == 2)
                            rarity2++;
                        else if (targetCardRarity == 3) rarity3++;
                    }
                }
            }

            groupRarityCount.Add((cardGroupName, rarity1, rarity2, rarity3));
        }

        var weight1 = (int)collectCardRarityDataList[0][cardRarityWeightIndex];
        var weight2 = (int)collectCardRarityDataList[1][cardRarityWeightIndex];
        var weight3 = (int)collectCardRarityDataList[2][cardRarityWeightIndex];

        var score1 = (int)collectCardRarityDataList[0][cardRarityScoreIndex];
        var score2 = (int)collectCardRarityDataList[1][cardRarityScoreIndex];
        var score3 = (int)collectCardRarityDataList[2][cardRarityScoreIndex];
        var newGroupRarityCount = new List<(string, int, int, int)>();
        var maxScore = (int)collectCardScoreDataLIst[0][cardScoreIndex];
        var countRarity = groupRarityCount.Count;
        var newRarity1 = 0;
        var newRarity2 = 0;
        var newRarity3 = 0;
        var simTimes = 100000;

        //模拟期望
        for (var i = 0; i < countRarity; i++)
        {
            newRarity1 += groupRarityCount[i].Item2;
            newRarity2 += groupRarityCount[i].Item3;
            newRarity3 += groupRarityCount[i].Item4;
            newGroupRarityCount.Add((groupRarityCount[i].Item1, newRarity1, newRarity2, newRarity3));

            if (groupRarityCount[i].Item2 == 0) weight1 = 0;
            if (groupRarityCount[i].Item3 == 0) weight2 = 0;
            if (groupRarityCount[i].Item4 == 0) weight3 = 0;
            //各自期望
            //var randCountSelf = RandCount(groupRarityCount,i, newGroupRarityCount, maxScore, score3, score2, score1, weight1, weight2, weight3, simTimes);
            //累积期望
            var randCountTotal = RandCount(groupRarityCount, i, newGroupRarityCount, maxScore, score3, score2, score1,
                weight1, weight2, weight3, simTimes);
            //Debug.Print("各自尝试次数：[" + randCountSelf.Item1 + "] ## " + "累积尝试次数：[" + randCountTotal.Item1 + "]");
            //collectCardGroup.Cells[i + 5, cardGroupIndex + 2].Value = randCountSelf.Item1;
            //collectCardGroup.Cells[i + 5, cardGroupIndex + 3].Value = randCountSelf.Item2;
            collectCardGroup.Cells[i + 5, cardGroupIndex + 4].Value = randCountTotal.Item1;
            collectCardGroup.Cells[i + 5, cardGroupIndex + 5].Value = randCountTotal.Item2;
            for (var a = 0; a < randCountTotal.Item3.Count; a++)
                collectCardGroup.Cells[a + 5, cardGroupIndex + 6 + i].Value = randCountTotal.Item3[a];
        }
        //计算各自期望--计算公式有问题
        //for (int i = 0; i < countRarity; i++)
        //{
        //    var rarity1Count = groupRarityCount[i].Item2;
        //    var rarity2Count = groupRarityCount[i].Item3;
        //    var rarity3Count = groupRarityCount[i].Item4;
        //    if (rarity1Count == 0)
        //    {
        //        weight1 = 0;
        //    }
        //    else if (rarity2Count == 0)
        //    {
        //        weight2 = 0;
        //    }
        //    else if (rarity3Count == 0)
        //    {
        //        weight3 = 0;
        //    }
        //    double ratio1 =(double) weight1 / (weight1 + weight2 + weight3);
        //    double ratio2 = (double)weight1 / (weight1 + weight2 + weight3);
        //    double ratio3 = (double)weight1 / (weight1 + weight2 + weight3);
        //    var eRarity = 1/ratio1+1/ratio2+1/ratio3-1/(ratio1+ratio2)-1/(ratio1+ratio3)-1/(ratio2+ratio3)+1/(ratio1+ratio2+ratio3);
        //    var eCard1 =PlusGG(rarity1Count);
        //    var eCard2 = PlusGG(rarity2Count);
        //    var eCard3 = PlusGG(rarity3Count);
        //    var eFinal = Math.Max(Math.Max(eCard1, eCard2), eCard3) * eRarity;
        //    Debug.Print("计算期望："+eFinal);
        //}
        ////计算累积期望
        //int PlusGG(int rarityCount)
        //{
        //    var sum =0;
        //    for (int i = 1; i <= rarityCount; i++)
        //    {
        //        sum += 1 / i;
        //    }
        //    var result = rarityCount * sum;
        //    return result;
        //}
    }

    private static Tuple<double, double, List<double>> RandCount(dynamic groupRarityCount, int i,
        dynamic newGroupRarityCount, int maxScore, int score3, int score2, int score1, int weight1, int weight2,
        int weight3, int simTimes)
    {
        var scoreGetTimes = 0;
        var otherRankCountTotal = new List<double>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
        var maxRankCountTotal = 0;

        var currentLimit1 = new List<(int, int)>();
        var currentLimit2 = new List<(int, int)>();
        var currentLimit3 = new List<(int, int)>();
        for (var k = 0; k < i + 1; k++)
        {
            var minValue1 = 0;
            var maxValue1 = newGroupRarityCount[k].Item2;
            var minValue2 = 0;
            var maxValue2 = newGroupRarityCount[k].Item3;
            var minValue3 = 0;
            var maxValue3 = newGroupRarityCount[k].Item4;
            if (k - 1 >= 0)
            {
                minValue1 = newGroupRarityCount[k - 1].Item2;
                minValue2 = newGroupRarityCount[k - 1].Item3;
                minValue3 = newGroupRarityCount[k - 1].Item4;
            }

            currentLimit1.Add((minValue1, maxValue1));
            currentLimit2.Add((minValue2, maxValue2));
            currentLimit3.Add((minValue3, maxValue3));
        }

        for (var s = 0; s < simTimes; s++)
        {
            var rarityRandom = new Random();
            var cardRandom = new Random();
            var cardList1 = new List<int>();
            var cardList2 = new List<int>();
            var cardList3 = new List<int>();
            var currentScore = 0;
            var otherRankCount = new List<double>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            var maxRankCount = 0;
            var currentCount1 = new List<int>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            var currentCount2 = new List<int>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            var currentCount3 = new List<int>() { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };


            while (cardList1.Count < newGroupRarityCount[i].Item2 || cardList2.Count < newGroupRarityCount[i].Item3 ||
                   cardList3.Count < newGroupRarityCount[i].Item4)
            {
                if (currentScore >= maxScore)
                {
                    var randMax = newGroupRarityCount[i].Item4;
                    var cardList = cardList3;
                    var score = score3;
                    if (randMax == 0)
                    {
                        randMax = newGroupRarityCount[i].Item3;
                        cardList = cardList2;
                        score = score2;
                        if (randMax == 0)
                        {
                            randMax = newGroupRarityCount[i].Item2;
                            cardList = cardList1;
                            score = score1;
                        }
                    }

                    int cardSeed = cardRandom.Next(1, randMax + 1);
                    if (cardList.Contains(cardSeed))
                    {
                        currentScore += score;
                    }
                    else
                    {
                        if (cardList == cardList1)
                        {
                            cardList1.Add(cardSeed);
                            for (var k = 0; k < i + 1; k++)
                                if (cardSeed <= currentLimit1[k].Item2 && cardSeed > currentLimit1[k].Item1)
                                    currentCount1[k]++;
                        }
                        else if (cardList == cardList2)
                        {
                            cardList2.Add(cardSeed);
                            for (var k = 0; k < i + 1; k++)
                                if (cardSeed <= currentLimit2[k].Item2 && cardSeed > currentLimit2[k].Item1)
                                    currentCount2[k]++;
                        }
                        else if (cardList == cardList3)
                        {
                            cardList3.Add(cardSeed);
                            for (var k = 0; k < i + 1; k++)
                                if (cardSeed <= currentLimit3[k].Item2 && cardSeed > currentLimit3[k].Item1)
                                    currentCount3[k]++;
                        }
                    }

                    currentScore -= maxScore;
                    scoreGetTimes++;
                }
                else
                {
                    var rairtySeed = rarityRandom.Next(1, weight1 + weight2 + weight3 + 1);
                    if (rairtySeed <= weight1 && weight1 != 0)
                    {
                        int cardSeed = cardRandom.Next(1, newGroupRarityCount[i].Item2 + 1);
                        if (cardList1.Contains(cardSeed))
                        {
                            currentScore += score1;
                        }
                        else
                        {
                            cardList1.Add(cardSeed);
                            for (var k = 0; k < i + 1; k++)
                                if (cardSeed <= currentLimit1[k].Item2 && cardSeed > currentLimit1[k].Item1)
                                    currentCount1[k]++;
                        }
                    }
                    else if (rairtySeed <= weight1 + weight2 && rairtySeed > weight1 && weight2 != 0)
                    {
                        int cardSeed = cardRandom.Next(1, newGroupRarityCount[i].Item3 + 1);
                        if (cardList2.Contains(cardSeed))
                        {
                            currentScore += score2;
                        }
                        else
                        {
                            cardList2.Add(cardSeed);
                            for (var k = 0; k < i + 1; k++)
                                if (cardSeed <= currentLimit2[k].Item2 && cardSeed > currentLimit2[k].Item1)
                                    currentCount2[k]++;
                        }
                    }
                    else if (rairtySeed <= weight1 + weight2 + weight3 && rairtySeed > weight1 + weight2 &&
                             weight3 != 0)
                    {
                        int cardSeed = cardRandom.Next(1, newGroupRarityCount[i].Item4 + 1);
                        if (cardList3.Contains(cardSeed))
                        {
                            currentScore += score3;
                        }
                        else
                        {
                            cardList3.Add(cardSeed);
                            for (var k = 0; k < i + 1; k++)
                                if (cardSeed <= currentLimit3[k].Item2 && cardSeed > currentLimit3[k].Item1)
                                    currentCount3[k]++;
                        }
                    }
                }

                maxRankCount++;
                for (var k = 0; k < i + 1; k++)
                    if (currentCount1[k] == groupRarityCount[k].Item2 &&
                        currentCount2[k] == groupRarityCount[k].Item3 &&
                        currentCount3[k] == groupRarityCount[k].Item4 && otherRankCount[k] == 0)
                        otherRankCount[k] = maxRankCount;
            }

            maxRankCountTotal += maxRankCount;
            for (var o = 0; o < otherRankCount.Count; o++) otherRankCountTotal[o] += otherRankCount[o];
        }

        otherRankCountTotal.RemoveAll(x => x == 0);
        for (var o = 0; o < otherRankCountTotal.Count; o++) otherRankCountTotal[o] /= simTimes;
        var resultValue = new Tuple<double, double, List<double>>((double)maxRankCountTotal / simTimes,
            (double)scoreGetTimes / simTimes, otherRankCountTotal);
        return resultValue;
    }
    //其他概率计算
}