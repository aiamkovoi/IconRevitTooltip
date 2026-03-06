using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OpenMcdf;

namespace IconRevitTooltip.Core
{
    public class RevitFileInfo
    {
        public string Version { get; set; } = "Unknown";
        public string Build { get; set; } = "Unknown";
        public string Worksharing { get; set; } = "Unknown";
        public string Locale { get; set; } = "Unknown";
        public string SaveCount { get; set; } = "Unknown";

        public static string GetVersionString(string path)
        {
            try
            {
                RevitFileInfo info = Parse(path);
                if (info != null && info.Version != "Unknown")
                {
                    string res = info.Version;
                    if (info.Worksharing != "Not enabled" && info.Worksharing != "Unknown")
                    {
                        res += $". Workshared! {info.Worksharing} file";
                    }
                    return res;
                }
            }
            catch { }
            return "UNKNOWN";
        }

        public static RevitFileInfo Parse(string path)
        {
            if (!File.Exists(path))
                return null;

            RevitFileInfo info = new RevitFileInfo();
            try
            {
                // First try direct file read
                int bytesToRead = 2 * 1024 * 1024;
                byte[] buffer;
                using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    int length = (int)Math.Min(bytesToRead, fs.Length);
                    buffer = new byte[bytesToRead];
                    fs.Read(buffer, 0, length);
                }
                string text = Encoding.UTF8.GetString(buffer);

                // Try to find Build before full parse
                Match buildMatch = Regex.Match(text, @"Build:\s*([^\r\n]*)");
                if (buildMatch.Success) info.Build = buildMatch.Groups[1].Value.Trim();

                Match match = Regex.Match(text, @"product-version>(\d+)<");
                if (match.Success)
                {
                    info.Version = match.Groups[1].Value;
                }
            }
            catch { }

            // Always try OLE as well to get more metadata
            try
            {
                using (CompoundFile cf = new CompoundFile(path, CFSUpdateMode.ReadOnly, CFSConfiguration.Default))
                {
                    CFStream stream = cf.RootStorage.GetStream("BasicFileInfo");
                    byte[] buffer0 = stream.GetData();
                    byte[] buffer = buffer0.Where(b => b != 0).ToArray();
                    string text = Encoding.UTF8.GetString(buffer);

                    Match mFormat = Regex.Match(text, @"Format:\s*(20\d{2})");
                    if (mFormat.Success) info.Version = mFormat.Groups[1].Value;
                    else
                    {
                        Match oldV = Regex.Match(text, @"Revit\s*(20\d{2})");
                        if (oldV.Success) info.Version = oldV.Groups[1].Value;
                    }

                    Match mBuild = Regex.Match(text, @"Build:\s*([^\r\n]*)");
                    if (mBuild.Success) info.Build = mBuild.Groups[1].Value.Trim();

                    Match mWs = Regex.Match(text, @"orksharing:\s*([^\r\n]*)");
                    if (mWs.Success) info.Worksharing = mWs.Groups[1].Value.Trim();

                    Match mLocale = Regex.Match(text, @"Locale[^:]*:\s*([^\r\n]*)");
                    if (mLocale.Success) info.Locale = mLocale.Groups[1].Value.Trim();

                    Match mSaves = Regex.Match(text, @"Increments[^:]*:\s*(\d+)");
                    if (mSaves.Success) info.SaveCount = mSaves.Groups[1].Value.Trim();
                }
            }
            catch { }

            return info;
        }
    }
}
