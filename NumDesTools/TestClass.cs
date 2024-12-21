using System.Threading.Tasks;
using Action = System.Action;

class Program
{
    public static async Task Main()
    {
        // 模拟输入数据
        object[,] rangeObj = {
            { "A1", "A2" },
            { "B1", "B2" }
        };
        object[,] rangeObjDef = {
            { "Default1", "Default2" },
            { "Default3", "Default4" }
        };
        string delimiter = "[,]";
        string ignoreValue = "";

        int iterations = 1000; // 调用次数

        // 测试同步方法
        Debug.Print("开始测试同步方法...");
        var syncTime = MeasureExecutionTime(() =>
        {
            for (int i = 0; i < iterations; i++)
            {
                CreatValueToArray(rangeObj, rangeObjDef, delimiter, ignoreValue);
            }
        });
        Debug.Print($"同步方法调用 {iterations} 次总耗时: {syncTime} ms");

        // 测试异步方法
        Debug.Print("开始测试异步方法...");
        var asyncTime = await MeasureExecutionTimeAsync(async () =>
        {
            for (int i = 0; i < iterations; i++)
            {
                await CreatValueToArrayAsync(rangeObj, rangeObjDef, delimiter, ignoreValue);
            }
        });
        Debug.Print($"异步方法调用 {iterations} 次总耗时: {asyncTime} ms");
    }

    // 测量同步方法的执行时间
    static long MeasureExecutionTime(Action syncMethod)
    {
        var stopwatch = Stopwatch.StartNew();
        syncMethod(); // 执行同步方法
        stopwatch.Stop();
        return stopwatch.ElapsedMilliseconds;
    }

    // 测量异步方法的执行时间
    static async Task<long> MeasureExecutionTimeAsync(Func<Task> asyncMethod)
    {
        var stopwatch = Stopwatch.StartNew();
        await asyncMethod(); // 等待异步方法完成
        stopwatch.Stop();
        return stopwatch.ElapsedMilliseconds;
    }

    // 同步版本的函数
    public static string CreatValueToArray(
        object[,] rangeObj,
        object[,] rangeObjDef,
        string delimiter,
        string ignoreValue)
    {
        var result = string.Empty;
        if (string.IsNullOrEmpty(delimiter))
        {
            delimiter = "[,]";
        }
        var delimiterList = delimiter.ToCharArray().Select(c => c.ToString()).ToArray();
        string startDelimiter;
        string midDelimiter;
        string endDelimiter;
        if (delimiterList.Length == 3)
        {
            startDelimiter = delimiterList[0];
            midDelimiter = delimiterList[1];
            endDelimiter = delimiterList[2];
        }
        else
        {
            startDelimiter = string.Empty;
            midDelimiter = delimiterList[0];
            endDelimiter = string.Empty;
        }
        var rows = rangeObj.GetLength(0);
        var cols = rangeObj.GetLength(1);
        for (var row = 0; row < rows; row++)
            for (var col = 0; col < cols; col++)
            {
                var item = rangeObj[row, col];
                if (item is ExcelEmpty || item.ToString() == ignoreValue || item is ExcelError) { }
                else
                {
                    if (!(rangeObjDef[0, 0] is ExcelMissing))
                    {
                        var itemDef = rangeObjDef[row, col];
                        result += itemDef + midDelimiter;
                    }
                    else
                    {
                        result += item + midDelimiter;
                    }
                }
            }

        if (!string.IsNullOrEmpty(result))
            result = startDelimiter + result.Substring(0, result.Length - midDelimiter.Length) + endDelimiter;
        return result;
    }

    // 异步版本的函数
    public static async Task<string> CreatValueToArrayAsync(
        object[,] rangeObj,
        object[,] rangeObjDef,
        string delimiter,
        string ignoreValue)
    {
        // 模拟异步操作
        return await Task.Run(() => CreatValueToArray(rangeObj, rangeObjDef, delimiter, ignoreValue));
    }
}
