
namespace EmailDeleter
{
    public class SimpleLogger
    {
        private readonly string _baseFileName;
        private readonly string _infoLogPath;

        public SimpleLogger(string baseFileName, string infoLogParg)
        {
            _baseFileName = baseFileName;
            _infoLogPath = infoLogParg;
        }

        private string GetLogFilePath()
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            return $"log-{date}.log";
        }
        private string GetInfoLogFilePath()
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            return Path.Combine(_infoLogPath, $"InfoLog-{date}.log");
        }

        public void LogInfo(string message)
        {
            var logFilePath = GetInfoLogFilePath();
            Log("INFO", message, logFilePath);
        }

        public void LogError(string message, Exception ex = null)
        {
            var errorMessage = ex == null ? message : $"{message}\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
            var logFilePath = GetLogFilePath();
            Log("ERROR", errorMessage, logFilePath);
        }

        private void Log(string logLevel, string message, string logFilePath)
        {
            var logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{logLevel}] {message}";
            try
            {
                

                // Ensure the directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(logFilePath));

                // Append the log to the file
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }
    }
}
