using Microsoft.Extensions.Logging;
using System.Text;
using System.Linq;
using System.Text.RegularExpressions;
using System;
using System.IO;

namespace TeamsStatus
{
    public enum MicrosoftTeamsStatus
    {
        Available,
        Busy,
        DoNotDisturb,
        Away,
        Offline,
        Unknown,
        OutOfOffice,
        InAMeeting
    }
    public interface IMicrosoftTeamsService
    {
        // in seconds
        int PoolingInterval { get; set; }

        //MicrosoftTeamsStatus GetCurrentStatus();
        String GetCurrentStatus();
    }
    public class MicrosoftTeamsService : IMicrosoftTeamsService
    {
        //private readonly ILogger<MicrosoftTeamsService> _logger;
        private MicrosoftTeamsStatus _lastStatus = MicrosoftTeamsStatus.Unknown;
        private DateTime _fileLastUpdated = DateTime.MinValue;

        public int PoolingInterval { get; set; }
        public string LogsFilePath { get; set; }

        public MicrosoftTeamsService() //ILogger<MicrosoftTeamsService> logger
        {
            //_logger = logger;

            this.PoolingInterval = Convert.ToInt32(5);
            this.LogsFilePath = Environment.ExpandEnvironmentVariables("%localappdata%\\Packages\\MSTeams_8wekyb3d8bbwe\\LocalCache\\Microsoft\\MSTeams\\Logs");
        }

        //public MicrosoftTeamsStatus GetCurrentStatus()
        public String GetCurrentStatus()
        {
            string pattern = @"MSTeams_*.log";
            var dirInfo = new DirectoryInfo(this.LogsFilePath);
            var fileInfo = (from f in dirInfo.GetFiles(pattern) orderby f.LastWriteTime descending select f).First();
            if (fileInfo.Exists)
            {
                var fileLastUpdated = fileInfo.LastWriteTime;
                if (fileLastUpdated > _fileLastUpdated)
                {
                    var lines = ReadLines(fileInfo.FullName);
                    foreach (var line in lines.Reverse())
                    {
                        var delFrom = @"GlyphBadge{""";
                        var delTo = @"""}, overlay";

                        if (line.Contains(delFrom) && line.Contains(delTo))
                        {
                            int posFrom = line.IndexOf(delFrom) + delFrom.Length;
                            var info2 = line.Substring(posFrom);
                            var status2 = info2.Split(@"""}").First();
                            int posTo = line.IndexOf(delTo);

                            if (true)
                            {
                                //var info = line.Substring(posFrom, posTo - posFrom);
                                var status = status2.Split(" -> ").Last();
                                var newStatus = _lastStatus;

                                switch (status)
                                {
                                    case "available":
                                        newStatus = MicrosoftTeamsStatus.Available;
                                        return "Green";
                                        break;
                                    case "away":
                                        newStatus = MicrosoftTeamsStatus.Away;
                                        return "Yellow";
                                        break;
                                    case "busy":
                                    case "onThePhone":
                                        newStatus = MicrosoftTeamsStatus.Busy;
                                        return "Red";
                                        break;
                                    case "doNotDisturb":
                                    case "presenting":
                                        newStatus = MicrosoftTeamsStatus.DoNotDisturb;
                                        return "Red";
                                        break;
                                    case "beRightBack":
                                        newStatus = MicrosoftTeamsStatus.Away;
                                        return "Yellow";
                                        break;
                                    case "offline":
                                        newStatus = MicrosoftTeamsStatus.Offline;
                                        return "Yellow";
                                        break;
                                    case "newActivity":
                                        // ignore this - happens where there is a new activity: Message, Like/Action, File Upload
                                        // this is not a real status change, just shows the bell in the icon
                                        break;
                                    case "inAMeeting":
                                        newStatus = MicrosoftTeamsStatus.InAMeeting;
                                        return "Red";
                                        break;
                                    default:
                                        //_logger.LogWarning($"MS Teams status unknown: {status}");
                                        newStatus = MicrosoftTeamsStatus.Unknown;
                                        return "NA";
                                        break;
                                }

                                if (newStatus != _lastStatus)
                                {
                                    _lastStatus = newStatus;
                                   // _logger.LogInformation($"MS Teams status set to {_lastStatus}");
                                }
                                break;
                            }
                        }
                    }
                    _fileLastUpdated = fileLastUpdated;
                }
            }

            // TO DO
            return _lastStatus.ToString();
        }

        private IEnumerable<string> ReadLines(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 0x1000, FileOptions.SequentialScan))
            using (var sr = new StreamReader(fs, Encoding.UTF8))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    yield return line;
                }
            }
        }
    }
}