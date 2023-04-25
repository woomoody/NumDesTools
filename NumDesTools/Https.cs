using ExcelDna.Integration;
using System;
using System.Diagnostics;
using System.Net;
using System.Text;

namespace NumDesTools;

public class SearchEngine
{
    //excel中实现搜索功能，会按照google、bing、baidu的顺序检测是否能ping，否则检查下一个网站
    [ExcelFunction(IsHidden = true)]
    public static string GoogleSearch(string query)
    {
        const string seachIndex1 = "/search?q=";
        const string google = "https://www.google.com";
        var url = google + seachIndex1 + query;
        //if (PingWebsite(google))
        //{
        //    url = google + seachIndex1 + query;
        //}
        //else if(PingWebsite(bingChina))
        //{
        //    url = bingChina + seachIndex1 + query;
        //}
        //else if (PingWebsite(bingInternational))
        //{
        //    url = bingInternational + seachIndex1 + query;
        //}
        //else
        //{
        //    url = baidu + seachIndex2 + query;
        //}
        var result = new StringBuilder();
        try
        {
            Process.Start(url);
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
        const string seachIndex1 = "/search?q=";
        const string bingInternational = "https://www.bing.com";
        var url = bingInternational + seachIndex1 + query;
        var result = new StringBuilder();
        try
        {
            Process.Start(url);
            result.Append("Search successful!");
        }
        catch (Exception ex)
        {
            result.Append("Search failed: " + ex.Message);
        }

        return result.ToString();
    }

    [ExcelFunction(IsHidden = true)]
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