using ExcelDna.Integration;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System;
using System.Xml.Serialization;

namespace NumDesTools;

class AddInWatcher : IDisposable
{
    // For every directory we watch, keep track of all the add-ins that have files in that directory
    public readonly Dictionary<string, WatchedDirectory> WatchedDirectories = new();
    HashSet<WatchedAddIn> _dirtyAddIns = new();
    readonly object _dirtyLock = new();

    public AddInWatcher(AddInReload config)
    {
        foreach (var addIn in config.WatchedAddIns)
        {
            foreach (var file in addIn.WatchedFiles)
            {
                var directory = Path.GetDirectoryName(file.Path);
                if (!WatchedDirectories.TryGetValue(directory ?? throw new InvalidOperationException(), out var wd))
                {
                    wd = new WatchedDirectory(directory, InvalidateAddIn);
                }
                wd.WatchAddIn(addIn);
            }
        }
    }

    // Called in the event handler - don't do slow work here.
    void InvalidateAddIn(WatchedAddIn watchedAddIn)
    {
        lock (_dirtyLock)
        {
            _dirtyAddIns.Add(watchedAddIn);
            ExcelAsyncUtil.QueueAsMacro(ReloadDirtyAddIns);
        }
    }

    // Running in macro context.
    void ReloadDirtyAddIns()
    {
        HashSet<WatchedAddIn> dirtyCopy;
        lock (_dirtyLock)
        {
            dirtyCopy = _dirtyAddIns;
            _dirtyAddIns = new HashSet<WatchedAddIn>();
        }
        foreach (var addIn in dirtyCopy)
        {
            ReloadAddIn(addIn.Path);
        }

        // Force a recalculate on open workbooks.
        XlCall.Excel(XlCall.xlcCalculateNow);
    }

    // Running in macro context.
    static void ReloadAddIn(string xllPath)
    {
        ExcelIntegration.RegisterXLL(xllPath);
    }

    public void Dispose()
    {
        foreach (var wd in WatchedDirectories.Values)
        {
            wd.Dispose();
        }
    }

    internal class WatchedDirectory : IDisposable
    {
        readonly FileSystemWatcher _directoryWatcher;
        readonly Dictionary<string, WatchedAddIn> _watchedFiles;
        readonly Action<WatchedAddIn> _invalidateAddIn;

        public WatchedDirectory(string path, Action<WatchedAddIn> invalidateAddIn)
        {
            _directoryWatcher = new FileSystemWatcher(path);
            _directoryWatcher.NotifyFilter = NotifyFilters.LastWrite;
            _directoryWatcher.Changed += DirectoryWatcher_Changed;
            _watchedFiles = new Dictionary<string, WatchedAddIn>(StringComparer.OrdinalIgnoreCase);
            _invalidateAddIn = invalidateAddIn;

            _directoryWatcher.EnableRaisingEvents = true;
        }

        public void WatchAddIn(WatchedAddIn addIn)
        {
            foreach (var file in addIn.WatchedFiles)
            {
                var fullPath = Path.GetFullPath(file.Path);
                _watchedFiles[fullPath] = addIn; // This only allows one add-in to watch a particular file.
            }
        }

        public void Dispose()
        {
            _directoryWatcher.Dispose();
        }

        void DirectoryWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            Debug.Assert(string.Equals(Path.GetFullPath(e.FullPath), e.FullPath, StringComparison.OrdinalIgnoreCase));

            if (_watchedFiles.TryGetValue(e.FullPath, out var addIn))
            {
                _invalidateAddIn(addIn);
            }
        }
    }
}
[Serializable]
[XmlType(AnonymousType = true)]
[XmlRoot(Namespace = "", IsNullable = false)]
public class AddInReload
{
    [XmlElement("WatchedAddIn", typeof(WatchedAddIn))]
    public List<WatchedAddIn> WatchedAddIns { get; set; }
}

[Serializable]
public class WatchedAddIn
{
    [XmlAttribute]
    public string Path { get; set; }
    [XmlElement("WatchedFile", typeof(WatchedFile))]
    public List<WatchedFile> WatchedFiles { get; set; }
}

[Serializable]
public class WatchedFile
{
    [XmlAttribute]
    public string Path { get; set; }
}