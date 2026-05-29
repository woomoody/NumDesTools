using System.Collections.Concurrent;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using MiniExcelLibs;
using NumDesTools.Config;
using OfficeOpenXml;
using DataTable = System.Data.DataTable;
using ExcelReference = ExcelDna.Integration.ExcelReference;

// ReSharper disable All

#pragma warning disable CA1416

namespace NumDesTools;

/// <summary>
/// 公共Excel功能类 — 职责拆分见同目录下的 partial 文件：
/// ExcelRw.cs（读写）/ ArrayConv.cs（数组转换）/ ExcelCheck.cs（校验）
/// </summary>
public static partial class PubMetToExcel { }
