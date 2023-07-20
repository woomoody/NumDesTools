//using System.Diagnostics;
//using SharpSvn;

//namespace NumDesTools;
//internal class SvnTools
//{
//    //static string path = @"C:\ProWork\trunk\Client\Assets\Resources\Table\Skill.txt";
//    //Revert并Update文件：后台
//    public static void UpFiles(string path, string logs)
//    {
//        var client = new SvnClient();
//        var comArg = new SvnCommitArgs
//        {
//            Depth = SvnDepth.Empty,
//            LogMessage = logs
//        };
//        //client.Revert(path);
//        //client.Update(path);
//        client.Commit(path, comArg, out _);
//    }

//    //Update文件夹：前台
//    public static void UpdateFiles(string path)
//    {
//        //SvnClient client = new SvnClient();
//        //client.Update(path);
//        Process.Start("TortoiseProc.exe", @"/command:update /path:" + path);
//    }

//    //获取文件Log
//    //public static void FileLogs()
//    //{
//    //    SvnClient client = new SvnClient();
//    //    SvnLogArgs args = new SvnLogArgs();
//    //    args.RetrieveAllProperties = false;//不检索所有属性
//    //    Collection<SvnLogEventArgs> status;
//    //    client.GetLog(path,args,out status);
//    //    var lognum = 0;
//    //    var logtext = "";
//    //    var lastlog = "";
//    //    foreach( var item in status)
//    //    {
//    //        if (lognum > 50)
//    //            break;
//    //        lognum += 1;
//    //        if(string.IsNullOrEmpty(item.LogMessage) || item.LogMessage == "" || lastlog == item.LogMessage)
//    //        {
//    //            continue;
//    //        }
//    //        logtext = item.Time + "=" + item.Author + ":" + item.LogMessage + "\n" + logtext;
//    //        lastlog = item.LogMessage;
//    //    }
//    //}
//    //提交文件:前台
//    public static void CommitFile(string path)
//    {
//        Process.Start("TortoiseProc.exe", @"/command:commit /path:" + path);
//        Process.Start("TortoiseProc.exe", @"/command:status /path:" + path);
//    }

//    //展示Log：前台
//    public static void FileLogs(string path)
//    {
//        Process.Start("TortoiseProc.exe", @"/command:log /path:" + path);
//    }

//    //与最近文件对比：前台
//    public static void DiffFile(string path)
//    {
//        Process.Start("TortoiseProc.exe", @"/command:diff /path:" + path);
//    }
//}

