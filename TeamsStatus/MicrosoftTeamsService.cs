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
        private Regex regex;
        public int PoolingInterval { get; set; }
        public string LogsFilePath { get; set; }


        public MicrosoftTeamsService() //ILogger<MicrosoftTeamsService> logger
        {
            //_logger = logger;
            this.regex = new Regex("SetBadge Setting badge: NumericBadge\\{\\d\\}, overlay: \\d items, ([A-Za-z]+( [A-Za-z]+)+)", RegexOptions.IgnoreCase);
            this.PoolingInterval = Convert.ToInt32(5);
            this.LogsFilePath = Environment.ExpandEnvironmentVariables("%localappdata%\\Packages\\MSTeams_8wekyb3d8bbwe\\LocalCache\\Microsoft\\MSTeams\\Logs"); //%localappdata%\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Logs
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
                        var status = "";
                        if (this.regex.IsMatch(line))
                        {
                            Match match = this.regex.Match(line);
                            status = match.Groups[2].Value.Trim(' ');

                            var newStatus = _lastStatus;

                            switch (status.ToLower())
                            {
                                case "available": //SetBadge Setting badge: NumericBadge{2}, overlay: 2 items, status Available
                                    newStatus = MicrosoftTeamsStatus.Available; //SetBadge Setting badge: NumericBadge{2}, overlay: 2 items, status Available
                                    return "Green";
                                    break;
                                case "away":
                                    newStatus = MicrosoftTeamsStatus.Away;
                                    return "Yellow";
                                    break;
                                case "busy": //SetBadge Setting badge: NumericBadge{2}, overlay: 2 items, status Busy
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
                            //}
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
