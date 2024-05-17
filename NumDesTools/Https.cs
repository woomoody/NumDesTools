using System.Net;
using System.Text;

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// Excel网页搜索引擎
/// </summary>
public class SearchEngine
{
    [ExcelFunction(IsHidden = true)]
    public static string GoogleSearch(string query)
    {
        var result = new StringBuilder();
        try
        {
            result.Append("Search successful!");
        }
        catch (Exception ex)
        {
            result.Append("Search failed: " + ex.Message);
        }

        return result.ToString();
    }

    [ExcelFunction(IsHidden = true)]
    public static string BingSearch(string query)
    {
        var result = new StringBuilder();
        try
        {
            result.Append("Search successful!");
        }
        catch (Exception ex)
        {
            result.Append("Search failed: " + ex.Message);
        }

        return result.ToString();
    }

    [ExcelFunction(IsHidden = true)]
    [Obsolete("Obsolete")]
    public static bool PingWebsite(string url)
    {
        bool isPass;
        try
        {
            var request = WebRequest.Create(url);
            request.GetResponse();
            isPass = true;
        }
        catch
        {
            isPass = false;
        }

        return isPass;
    }
}