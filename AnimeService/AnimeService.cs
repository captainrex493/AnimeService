using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace AnimeService
{
    public enum ServiceState
    {
        SERVICE_STOPPED = 0x00000001,
        SERVICE_START_PENDING = 0x00000002,
        SERVICE_STOP_PENDING = 0x00000003,
        SERVICE_RUNNING = 0x00000004,
        SERVICE_CONTINUE_PENDING = 0x00000005,
        SERVICE_PAUSE_PENDING = 0x00000006,
        SERVICE_PAUSED = 0x00000007,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct ServiceStatus
    {
        public int dwServiceType;
        public ServiceState dwCurrentState;
        public int dwControlsAccepted;
        public int dwWin32ExitCode;
        public int dwServiceSpecificExitCode;
        public int dwCheckPoint;
        public int dwWaitHint;
    }

    public partial class AnimeService : ServiceBase
    {
        EventLog eventLog;
        FileSystemWatcher fileSystemWatcher;


        public AnimeService()
        {
            InitializeComponent();


            eventLog = new System.Diagnostics.EventLog();
            if (!System.Diagnostics.EventLog.SourceExists("AnimeSource"))
                {
                    System.Diagnostics.EventLog.CreateEventSource(
                        "AnimeSource", "AnimeLog");
                }
            eventLog.Source = "AnimeSource";
            eventLog.Log = "AnimeLog";

            fileSystemWatcher = new FileSystemWatcher();
            fileSystemWatcher.Path = "C:\\Users\\keval\\Downloads";

            fileSystemWatcher.NotifyFilter = NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.FileName
                                 | NotifyFilters.DirectoryName;

            fileSystemWatcher.Filter = "*.mkv";

            fileSystemWatcher.Created += OnCreated;

            fileSystemWatcher.EnableRaisingEvents = true;
            eventLog.WriteEntry("Finished service initialization", EventLogEntryType.Information);
        }

        private void OnCreated(object source, FileSystemEventArgs e)
        {
            eventLog.WriteEntry("Event raised for " + e.Name, EventLogEntryType.Information);
            int count = 0;

            while (!IsFileReady(e.FullPath))
            {

                if (count > 600)
                {
                    eventLog.WriteEntry(e.Name + " was in use and timed out", EventLogEntryType.Error, 403);
                    return;
                }
                else if (!File.Exists(e.FullPath))
                {
                    eventLog.WriteEntry(e.Name + " no longer exists", EventLogEntryType.Error, 410);
                    return;
                }

                System.Threading.Thread.Sleep(1000);
                count++;
            }

            DirectoryInfo animeFolder = new DirectoryInfo("C:\\Users\\keval\\anime");
            DirectoryInfo bestMatch = new DirectoryInfo("C:\\Users\\keval\\anime\\");
            int bestRatio = 0;
            foreach  (DirectoryInfo folder in animeFolder.GetDirectories())
            {
                int ratio = FuzzySharp.Fuzz.PartialRatio(e.Name, folder.Name);
                eventLog.WriteEntry("Ratio of " + folder.Name + " and " + e.Name + "is: \t " + ratio, EventLogEntryType.Information);
                if (ratio > bestRatio)
                {
                    bestRatio = ratio;
                    bestMatch = folder;
                }
                
            }
            File.Move(e.FullPath, bestMatch.FullName + "\\" + e.Name);
            eventLog.WriteEntry("Moved " + e.Name + " to " + bestMatch.FullName + e.Name, EventLogEntryType.Information);

        }

        protected override void OnStart(string[] args)
        {
        }

        protected override void OnStop()
        {
        }

        public static bool IsFileReady(string filename)
        {
            // If the file can be opened for exclusive access it means that the file
            // is no longer locked by another process.
            try
            {
                using (FileStream inputStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None))
                    return inputStream.Length > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern bool SetServiceStatus(System.IntPtr handle, ref ServiceStatus serviceStatus);
    }
}
