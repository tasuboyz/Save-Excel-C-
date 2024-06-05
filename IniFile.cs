using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Interfaccia_Database_sql_C_
{
    public class IniFile
    {
        private string _filePath;

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        public IniFile(string filePath)
        {
            _filePath = filePath;
        }

        public void Write(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, _filePath);
        }

        public string Read(string section, string key)
        {
            StringBuilder sb = new StringBuilder(255);
            int i = GetPrivateProfileString(section, key, "", sb, 255, _filePath);
            return sb.ToString();
        }
        public void CheckAndDeleteIniFile()
        {
            string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string iniFilePath = Path.Combine(appDirectory, "PathConfig.ini");

            if (File.Exists(iniFilePath))
            {
                DateTime lastModified = File.GetLastWriteTime(iniFilePath);
                DateTime currentDate = DateTime.Now;

                // Confronta le date
                if (currentDate.Subtract(lastModified).TotalMinutes >= 30)
                {
                    // Leggi tutte le righe del file
                    var lines = File.ReadAllLines(iniFilePath).ToList();

                    // Rimuovi la riga corrispondente a ExcelPath
                    lines.RemoveAll(line => line.Contains("ExcelPath"));

                    // Scrivi le righe rimanenti nel file
                    File.WriteAllLines(iniFilePath, lines);
                }
            }
        }
    }
    public class DirectoryPaths
    {
        private string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
        //private string newDirectory = ".";

        public string GetConfigIniPath()
        {
            return Path.Combine(appDirectory, "config.ini");
        }
        public string GetExcelIniPath()
        {
            return Path.Combine(appDirectory, "PathConfig.ini");
        }
        public string GetLogFilePath()
        {
            return Path.Combine(appDirectory, "log.log");
        }
    }
    public class Logger
    {
        private string logFilePath;
        public Logger(string filePath)
        {
            this.logFilePath = filePath;
        }

        public void Log(string message)
        {
            using (StreamWriter writer = new StreamWriter(logFilePath, false)) //false per indicare che i log non vengono sovrascritti
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
