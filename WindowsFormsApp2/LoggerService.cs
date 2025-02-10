using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace WindowsFormsApp2
{
    public static class LoggerService
    {
        private static readonly string machineName = Environment.MachineName; // Get PC Name
        private static readonly string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SkyCrawler");
        private static readonly string logDirectory = Path.Combine(appDataPath, "logs");
        private static readonly string logFilePath = Path.Combine(logDirectory, $"log-{machineName}-.txt"); // Unique log file per PC

        static LoggerService()
        {
            try
            {
                // Ensure the logs folder exists inside AppData
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // Configure Serilog with PC Name
                Log.Logger = new LoggerConfiguration()
                    .Enrich.WithProperty("MachineName", machineName) // Add PC Name to each log entry
                    .WriteTo.Console()
                    .WriteTo.File(logFilePath,
                        rollingInterval: RollingInterval.Day,  // New log file each day
                        retainedFileCountLimit: 7,             // Keep logs for 7 days
                        fileSizeLimitBytes: 10_000_000,        // 10MB per file max
                        rollOnFileSizeLimit: true)            // Create a new file if exceeded
                    .CreateLogger();

                Log.Information("Logger initialized on {MachineName}. Logs will be saved in: {Path}", machineName, logDirectory);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error initializing logger: " + ex.Message);
            }
        }

        public static void Info(string message)
        {
            Log.Information("[{MachineName}] {Message}", machineName, message);
        }

        public static void Warning(string message)
        {
            Log.Warning("[{MachineName}] {Message}", machineName, message);
        }

        public static void Error(string message, Exception ex = null)
        {
            if (ex == null)
                Log.Error("[{MachineName}] {Message}", machineName, message);
            else
                Log.Error(ex, "[{MachineName}] {Message}", machineName, message);
        }

        public static void Close()
        {
            Log.CloseAndFlush();
        }
    }
}
