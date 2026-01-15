using System;
using System.IO;
using System.Text;
using System.Xml;
using log4net;

namespace Spreadsheet.Handler
{
    internal static class Logger
    {
        private const int DaysDifference = 5;
        private const string DebugLevel = "ERROR";
        private const string AppenderName = "EmpowerImportAppender";
        private const string LoggerName = "EmpowerImport";

        static Logger()
        {
            InitLogger();
            Log = _log;
        }

        private static ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static ILog Log { get; private set; }


        private static void InitLogger()
        {
            // Do some cleanup by removing the now "old" log file
            string defaultPath = GetDefaultPath();
            ClearPathLogs(defaultPath);

            string loggerPath = GetNextLoggerFile(GetDefaultPath());

            string loggerConfig = LoggerDefaultConfig(loggerPath);

            XmlDocument doc = new XmlDocument();
            try
            {
                doc.LoadXml(loggerConfig);
                log4net.Config.XmlConfigurator.Configure(doc.DocumentElement);

                _log = LogManager.GetLogger(LoggerName);
                LogMessage("Logger configured successfully.", log4net.Core.Level.Info);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to setup the logger! " + ex.Message);
            }
        }

        /// <summary>
        /// Assume the config file is located in the same directory as this Assembly
        /// </summary>
        /// <returns></returns>
        private static string GetDefaultPath()
        {
            string appPath = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
            string strPath = Path.GetDirectoryName(appPath);
            if (strPath.StartsWith(@"file:\"))
            {
                strPath = strPath.Substring(6);
            }
            return (strPath);
        }


        private static string GetNextLoggerFile(string defaultPath)
        {
            int counter = 1;
            while (File.Exists(Path.Combine(defaultPath, string.Format("Logger{0}.log", counter))))
            {
                counter++;
            }
            string strPath = Path.Combine(defaultPath, string.Format("Logger{0}.log", counter));

            return (strPath);
        }


        /// <summary>
        /// Provides the default logger configuration xml.
        /// </summary>
        /// <returns>The configuration xml string.</returns>
        private static string LoggerDefaultConfig(string loggerLogFilePath)
        {
            StringBuilder xml = new StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
            xml.Append("\r\n<log4net>");
            xml.Append("\r\n<appender name=\"" + AppenderName + "\" type=\"log4net.Appender.FileAppender\">");
            xml.Append("\r\n<file value=\"" + loggerLogFilePath + "\" />");
            xml.Append("\r\n<appendToFile value=\"true\" />");
            xml.Append("\r\n<layout type=\"log4net.Layout.SimpleLayout\" />");
            xml.Append("\r\n</appender>");
            xml.Append("\r\n<logger name=\"" + LoggerName + "\">");
            xml.Append("\r\n<level value=\"" + DebugLevel + "\" />");
            xml.Append("\r\n<appender-ref ref=\"" + AppenderName + "\" />");
            xml.Append("\r\n</logger>");
            xml.Append("\r\n");
            xml.Append("\r\n</log4net>");

            return xml.ToString();
        }


        public static void LogMessage(string msg, log4net.Core.Level level)
        {
            if (_log == null) return;

            msg = DateTime.Now.ToString("u") + ": " + msg;

            if (level == log4net.Core.Level.Error && _log.IsErrorEnabled)
            {
                _log.Error(msg);
            }
            else if (level == log4net.Core.Level.Fatal && _log.IsFatalEnabled)
            {
                _log.Fatal(msg);
            }
            else if (level == log4net.Core.Level.Warn && _log.IsWarnEnabled)
            {
                _log.Warn(msg);
            }
            else if (level == log4net.Core.Level.Info && _log.IsInfoEnabled)
            {
                _log.Info(msg);
            }
            else if (level == log4net.Core.Level.Debug && _log.IsDebugEnabled)
            {
                _log.Debug(msg);
            }
            else if (level == log4net.Core.Level.Error && _log.IsErrorEnabled)
            {
                _log.Error(msg);
            }
        }


        private static void ClearPathLogs(string path)
        {
            string[] filePaths;
            try
            {
                // This can fail if access to the path is denied or if path does not exist
                filePaths = Directory.GetFiles(@path, "*.log");
            }
            catch (Exception)
            {
                return;
            }

            foreach (string filePath in filePaths)
            {
                if (File.Exists(filePath))
                {
                    DateTime fileDateTime = File.GetCreationTime(filePath);
                    if (DateTime.Now.AddDays(-DaysDifference) > fileDateTime)
                    {
                        try
                        {
                            File.Delete(filePath);
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
            }
        }

    }
}