using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using SharpShell.Attributes;
using SharpShell.SharpInfoTipHandler;
using IconRevitTooltip.Core;

namespace IconRevitTooltip.ShellExtension
{
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".rvt", ".rfa")]
    public class RevitInfoTipExtension : SharpInfoTipHandler
    {
        static RevitInfoTipExtension()
        {
            AppDomain.CurrentDomain.AssemblyResolve += (s, e) =>
            {
#if DEBUG
                string logPath = Path.Combine(Path.GetTempPath(), "RevitTooltipLog.txt");
                File.AppendAllText(logPath, $"[Resolve] Trying to load {e.Name}\n");
#endif
                string folder = Path.GetDirectoryName(typeof(RevitInfoTipExtension).Assembly.Location);
                string path = Path.Combine(folder, new System.Reflection.AssemblyName(e.Name).Name + ".dll");
                return File.Exists(path) ? System.Reflection.Assembly.LoadFrom(path) : null;
            };
        }

        protected override string GetInfo(RequestedInfoType infoType, bool singleLine)
        {
#if DEBUG
            string logPath = Path.Combine(Path.GetTempPath(), "RevitTooltipLog.txt");
            File.AppendAllText(logPath, $"[GetInfo] Called for {infoType}, File: {SelectedItemPath}\n");
#endif

            if (infoType == RequestedInfoType.InfoTip)
            {
                try
                {
                    if (string.IsNullOrEmpty(SelectedItemPath) || !File.Exists(SelectedItemPath))
                    {
#if DEBUG
                        File.AppendAllText(logPath, "[GetInfo] Path empty or invalid.\n");
#endif
                        return string.Empty;
                    }

#if DEBUG
                    File.AppendAllText(logPath, "[GetInfo] Calling RevitFileInfo.Parse...\n");
#endif
                    var info = RevitFileInfo.Parse(SelectedItemPath);
                    if (info == null)
                    {
#if DEBUG
                        File.AppendAllText(logPath, "[GetInfo] Parse returned null.\n");
#endif
                        return string.Empty;
                    }

#if DEBUG
                    File.AppendAllText(logPath, $"[GetInfo] Parse success! V={info.Version} B={info.Build} W={info.Worksharing} L={info.Locale} S={info.SaveCount}\n");
#endif
                    var builder = new StringBuilder();

                    if (!string.IsNullOrEmpty(info.Version) && info.Version != "Unknown")
                        builder.AppendLine($"Revit Version: {info.Version}");
                    
                    if (!string.IsNullOrEmpty(info.Build) && info.Build != "Unknown")
                        builder.AppendLine($"Build: {info.Build}");
                    
                    if (!string.IsNullOrEmpty(info.Worksharing) && info.Worksharing != "Unknown")
                    {
                        string wsStatus = info.Worksharing.Contains("Not") ? "No" : "Yes (" + info.Worksharing + ")";
                        builder.AppendLine($"Worksharing: {wsStatus}");
                    }

                    if (!string.IsNullOrEmpty(info.Locale) && info.Locale != "Unknown")
                        builder.AppendLine($"Locale: {info.Locale}");

                    if (!string.IsNullOrEmpty(info.SaveCount) && info.SaveCount != "Unknown")
                        builder.AppendLine($"Saves: {info.SaveCount}");

                    string result = builder.ToString().TrimEnd();

                    if (singleLine)
                    {
                        return result.Replace(Environment.NewLine, " | ");
                    }

#if DEBUG
                    File.AppendAllText(logPath, $"[GetInfo] Returning: [{result}]\n");
#endif
                    return result;
                }
                catch (Exception ex)
                {
#if DEBUG
                    File.AppendAllText(logPath, $"[GetInfo] ERROR: {ex.Message} {ex.StackTrace}\n");
#endif
                    return $"Error reading Revit metadata: {ex.Message}";
                }
            }

            return string.Empty;
        }
    }
}
