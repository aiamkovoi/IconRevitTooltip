using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using OpenMcdf;
using System.Reflection;

[assembly: AssemblyVersion("1.0.*")]
namespace IconRevitTooltip
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string msg = string.Empty;
            string caption = string.Empty;

            if (args.Length < 2)
            {
                string filePath = string.Empty;
                if (args.Length == 0)
                {
                    OpenFileDialog ofd = new OpenFileDialog();
                    string filter = "Revit files (*.rvt;*.rfa)|*.rvt;*.rfa|All files (*.*)|*.*";
                    ofd.Filter = filter;
                    ofd.Multiselect = false;
                    ofd.Title = "Version " + Assembly.GetExecutingAssembly().GetName().Version.ToString();
                    if (ofd.ShowDialog() != DialogResult.OK)
                        return;

                    filePath = ofd.FileName;
                }
                else if (args.Length == 1)
                {
                    filePath = args[0];
                }
                string version = OpenFileAndGetVersion(filePath);
                caption = Path.GetFileName(filePath);
                msg = $"Revit {version}";

            }
            else
            {
                List<string> lines = new List<string>();
                foreach (string s in args)
                {
                    string version = OpenFileAndGetVersion(s);
                    string filename = Path.GetFileName(s);
                    lines.Add($"{filename}: \tRevit {version}");
                }
                msg = string.Join(Environment.NewLine, lines);
            }
            MessageBox.Show(msg, caption);
        }

        private static string OpenFileAndGetVersion(string path)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException("File not found: " + path);
            }
            string value = IconRevitTooltip.Core.RevitFileInfo.GetVersionString(path);
            if (value == "UNKNOWN" || string.IsNullOrEmpty(value))
            {
                string errMsg = "Failed to determine the Revit version: " + path;
                MessageBox.Show(errMsg);
                throw new Exception(errMsg);
            }
            return value;
        }
    }
}
