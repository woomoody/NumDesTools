using ExcelDna.Integration;

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

        var collectCardGroupData = PubMetToExcel.ExcelDataToListBySelf(collectCardGroup, 5, 1, 2,1);
        var collectCardGroupTitle = collectCardGroupData.Item1;
        var collectCardGroupDataList = collectCardGroupData.Item2;
        var collectCardInfoData = PubMetToExcel.ExcelDataToListBySelf(collectCardInfo, 5, 1, 2,1);
        var collectCardInfoTitle = collectCardInfoData.Item1;
        var collectCardInfoDataList = collectCardInfoData.Item2;
        var collectCardRarityData = PubMetToExcel.ExcelDataToListBySelf(collectCardRarity, 5, 1, 2,1);
        var collectCardRarityTitle = collectCardRarityData.Item1;
        var collectCardRarityDataList = collectCardRarityData.Item2;
        var collectCardScoreData = PubMetToExcel.ExcelDataToListBySelf(collectCardScore, 5, 1, 2,1);
        var collectCardScoreTitle = collectCardScoreData.Item1;
        var collectCardScoreDataLIst = collectCardScoreData.Item2;
        //分拆出每个卡组的稀有度构成


    }

    //创建卡牌序列

    //随机抽卡，记录重复while循环，返回值
}