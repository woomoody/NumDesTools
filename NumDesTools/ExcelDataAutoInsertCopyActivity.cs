#pragma warning disable CA1416

namespace NumDesTools;

public static class ExcelDataAutoInsertCopyActivity
{
    public static void RightClickCloneData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件
        var wkPath = AppServices.App.ActiveWorkbook.Path;
        var excelNames = new List<string>()
        {
            "RechargeAmazon.xlsx",
            "RechargeAptoide.xlsx",
            "RechargeGlobalOfficial.xlsx",
            "RechargeSamsung.xlsx",
            "RechargeIOS.xlsx"
        };
        var defaultValues = new Dictionary<string, List<string>>()
        {
            {
                "thirdProductID",
                new List<string>() { "com.mergeland.alices.adventure_diamond_", "price_Num" }
            },
        };

        var replaceValues = new Dictionary<string, Dictionary<string, List<string>>>()
        {
            {
                "RechargeIOS.xlsx",
                new Dictionary<string, List<string>>()
                {
                    {
                        "productID",
                        new List<string>() { "mergeland.alices.adventure", "casualgame.type.pipe" }
                    },
                    {
                        "productID_test",
                        new List<string>() { "mergeland.alices.adventure", "casualgame.type.pipe" }
                    }
                }
            }
        };

        ExcelDataSyncHelper.SyncSelectedRows(
            targetPath: wkPath,
            targetFileNames: excelNames,
            defaultValues: defaultValues,
            replaceValues: replaceValues
        );
    }

    public static void RightClickCloneAllData(CommandBarButton ctrl, ref bool cancelDefault)
    {
        cancelDefault = true; // 阻止默认事件
        var wkPath = AppServices.App.ActiveWorkbook.Path;
        var excelNames = new List<string>()
        {
            "RechargeAmazon.xlsx",
            "RechargeAptoide.xlsx",
            "RechargeGlobalOfficial.xlsx",
            "RechargeSamsung.xlsx",
            "RechargeIOS.xlsx"
        };
        var defaultValues = new Dictionary<string, List<string>>()
        {
            {
                "thirdProductID",
                new List<string>() { "com.mergeland.alices.adventure_diamond_", "price_Num" }
            },
        };

        var replaceValues = new Dictionary<string, Dictionary<string, List<string>>>()
        {
            {
                "RechargeIOS.xlsx",
                new Dictionary<string, List<string>>()
                {
                    {
                        "productID",
                        new List<string>() { "mergeland.alices.adventure", "casualgame.type.pipe" }
                    },
                    {
                        "productID_test",
                        new List<string>() { "mergeland.alices.adventure", "casualgame.type.pipe" }
                    }
                }
            }
        };

        ExcelDataSyncHelper.SyncAllRows(
            targetPath: wkPath,
            targetFileNames: excelNames,
            sourceFileName: "RechargeGP.xlsx",
            defaultValues: defaultValues,
            replaceValues: replaceValues
        );
    }
}
