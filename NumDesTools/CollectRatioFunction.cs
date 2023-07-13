using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace NumDesTools;

public class CollectRatioFunction
{
    private static readonly dynamic App = ExcelDnaUtil.Application;
    public static void CacCardCollect()
    {
        //获取表格源数据
        var workBook = App.ActiveWorkbook;
        var excelPath = workBook.Path;
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
        for (int i = 0; i < groupCount; i++)
        {
            var cardGroupStr = collectCardGroupDataList[i][cardGroupIndex];
            var cardGroupName = collectCardGroupDataList[i][cardGroupNameIndex];
            //拆ID，查ID，获取各个品质的个数
            var reg = "\\d+";
            var matches = Regex.Matches(cardGroupStr, reg);
            int rarity1 = 0;
            int rarity2 = 0;
            int rarity3 = 0;
            foreach (var match in matches)
            {
                var sourceCardId = match.Value;
                for (int j = 0; j < cardCount; j++)
                {
                    var targetCardId = collectCardInfoDataList[j][cardInfoIdIndex].ToString();
                    if (targetCardId == sourceCardId)
                    {
                        var targetCardRarity = collectCardInfoDataList[j][cardInfoRarityIndex];
                        if (targetCardRarity == 1)
                        {
                            rarity1++;
                        }
                        else if (targetCardRarity == 2)
                        {
                            rarity2++;
                        }
                        else if (targetCardRarity == 3)
                        {
                            rarity3++;
                        }
                    }
                }
            }
            groupRarityCount.Add((cardGroupName, rarity1, rarity2, rarity3));
        }

        int weight1 = (int)collectCardRarityDataList[0][cardRarityWeightIndex];
        int weight2 = (int)collectCardRarityDataList[1][cardRarityWeightIndex];
        int weight3 = (int)collectCardRarityDataList[2][cardRarityWeightIndex];

        int score1 = (int)collectCardRarityDataList[0][cardRarityScoreIndex];
        int score2 = (int)collectCardRarityDataList[1][cardRarityScoreIndex];
        int score3 = (int)collectCardRarityDataList[2][cardRarityScoreIndex];
        var newGroupRarityCount = new List<(string, int, int, int)>();
        int maxScore = (int)collectCardScoreDataLIst[0][cardScoreIndex];
        var countRarity = groupRarityCount.Count;
        int newRarity1 = 0;
        int newRarity2 = 0;
        int newRarity3 = 0;
        int simuTimes = 1000000;

        //模拟期望
        for (int i = 0; i < countRarity; i++)
        {
            newRarity1 += groupRarityCount[i].Item2;
            newRarity2 += groupRarityCount[i].Item3;
            newRarity3 += groupRarityCount[i].Item4;
            newGroupRarityCount.Add((groupRarityCount[i].Item1, newRarity1, newRarity2, newRarity3));

            var currentScore = 0;
            if (groupRarityCount[i].Item2 == 0)
            {
                weight1 = 0;
            }
            else if (groupRarityCount[i].Item3 == 0)
            {
                weight2 = 0;
            }
            else if (groupRarityCount[i].Item4 == 0)
            {
                weight3 = 0;
            }
            //各自期望
            var randCountSelf = RandCount(groupRarityCount[i].Item2, groupRarityCount[i].Item3, groupRarityCount[i].Item4, currentScore, maxScore, score3, score2, score1, weight1, weight2, weight3, simuTimes);
            //累积期望
            var randCountTotal = RandCount(newGroupRarityCount[i].Item2, newGroupRarityCount[i].Item3, newGroupRarityCount[i].Item4, currentScore, maxScore, score3, score2, score1, weight1, weight2, weight3, simuTimes);
            Debug.Print("各自尝试次数：[" + randCountSelf + "] ## " + "累积尝试次数：[" + randCountTotal + "]");
            collectCardGroup.Cells[i + 5, cardGroupIndex + 2].Value = randCountSelf;
            collectCardGroup.Cells[i + 5, cardGroupIndex + 3].Value = randCountTotal;
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

    private static double RandCount(dynamic rarityCount1, dynamic rarityCount2, dynamic rarityCount3, int currentScore, int maxScore, int score3, int score2, int score1, int weight1, int weight2, int weight3, int simuTimes)
    {
        var simuCount = 0;
        for (int s = 0; s < simuTimes; s++)
        {
            var rarityRandom = new Random();
            var cardRandom = new Random();
            var cardList1 = new List<int>();
            var cardList2 = new List<int>();
            var cardList3 = new List<int>();
            var randCount = 0;
            while (cardList1.Count < rarityCount1 || cardList2.Count < rarityCount2 ||
                   cardList3.Count < rarityCount3)
            {
                if (currentScore >= maxScore)
                {
                    var randMax = rarityCount3;
                    var cardList = cardList3;
                    var score = score3;
                    if (randMax == 0)
                    {
                        randMax = rarityCount2;
                        cardList = cardList2;
                        score = score2;
                        if (randMax == 0)
                        {
                            randMax = rarityCount1;
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
                            //Debug.Print("品质1：" + cardSeed+"分："+currentScore);
                        }
                        else if (cardList == cardList2)
                        {
                            cardList2.Add(cardSeed);
                            //Debug.Print("品质1：" + cardSeed + "分：" + currentScore);
                        }
                        else if (cardList == cardList3)
                        {
                            cardList3.Add(cardSeed);
                            //Debug.Print("品质1：" + cardSeed + "分：" + currentScore);
                        }
                    }
                    currentScore -= maxScore;
                }
                else
                {
                    int rairtySeed = rarityRandom.Next(1, weight1 + weight2 + weight3 + 1);
                    //Debug.Print("品质seed：" +rairtySeed);
                    if (rairtySeed <= weight1 && weight1 != 0)
                    {
                        int cardSeed = cardRandom.Next(1, rarityCount1 + 1);
                        if (cardList1.Contains(cardSeed))
                        {
                            currentScore += score1;
                        }
                        else
                        {
                            cardList1.Add(cardSeed);
                        }
                        //Debug.Print("品质1："+cardSeed);
                    }
                    else if (rairtySeed <= weight1 + weight2 && rairtySeed > weight1 && weight2 != 0)
                    {
                        int cardSeed = cardRandom.Next(1, rarityCount2 + 1);
                        if (cardList2.Contains(cardSeed))
                        {
                            currentScore += score2;
                        }
                        else
                        {
                            cardList2.Add(cardSeed);
                        }
                        //Debug.Print("品质2：" + cardSeed);
                    }
                    else if (rairtySeed <= weight1 + weight2 + weight3 && rairtySeed > weight1 + weight2 &&
                             weight3 != 0)
                    {
                        int cardSeed = cardRandom.Next(1, rarityCount3 + 1);
                        if (cardList3.Contains(cardSeed))
                        {
                            currentScore += score3;
                        }
                        else
                        {
                            cardList3.Add(cardSeed);
                        }
                        //Debug.Print("品质3：" + cardSeed);
                    }
                }
                randCount++;
            }
            simuCount += randCount;
        }
        // ReSharper disable once PossibleLossOfFraction
        return (double)simuCount / simuTimes;
    }
}