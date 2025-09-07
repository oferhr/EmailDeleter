
namespace EmailDeleter
{
    public class SimpleLogger
    {
        private readonly string _baseFileName;
        private readonly string _infoLogPath;
        private readonly bool _enableDebugLogging;

        public SimpleLogger(string baseFileName, string infoLogParg, bool enableDebugLogging = false)
        {
            _baseFileName = baseFileName;
            _infoLogPath = infoLogParg;
            _enableDebugLogging = enableDebugLogging;
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

        private string GetDebugLogFilePath()
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            return Path.Combine(_infoLogPath, $"DebugLog-{date}.log");
        }

        public void LogInfo(string message)
        {
            var logFilePath = GetInfoLogFilePath();
            Log("INFO", message, logFilePath);
            
            // When debug logging is enabled, also log info messages to debug log
            if (_enableDebugLogging)
            {
                var debugLogFilePath = GetDebugLogFilePath();
                Log("INFO", message, debugLogFilePath);
            }
        }

        public void LogWarning(string message)
        {
            var logFilePath = GetInfoLogFilePath();
            Log("WARNING", message, logFilePath);
            
            // When debug logging is enabled, also log warning messages to debug log
            if (_enableDebugLogging)
            {
                var debugLogFilePath = GetDebugLogFilePath();
                Log("WARNING", message, debugLogFilePath);
            }
        }

        public void LogDebug(string message)
        {
            if (_enableDebugLogging)
            {
                var logFilePath = GetDebugLogFilePath();
                Log("DEBUG", message, logFilePath);
            }
        }

        public void LogError(string message, Exception? ex = null)
        {
            var errorMessage = ex == null ? message : $"{message}\nException: {ex.Message}\nStackTrace: {ex.StackTrace}";
            var logFilePath = GetLogFilePath();
            Log("ERROR", errorMessage, logFilePath);
            
            // When debug logging is enabled, also log error messages to debug log
            if (_enableDebugLogging)
            {
                var debugLogFilePath = GetDebugLogFilePath();
                Log("ERROR", errorMessage, debugLogFilePath);
            }
        }

        public void LogPerformance(string operation, TimeSpan duration, string additionalInfo = "")
        {
            var message = $"Performance: {operation} took {duration.TotalMilliseconds:F2}ms {additionalInfo}";
            
            // Performance logs only go to debug logger when debug logging is enabled
            if (_enableDebugLogging)
            {
                var logFilePath = GetDebugLogFilePath();
                Log("PERFORMANCE", message, logFilePath);
            }
        }

        public void LogBatchOperation(string operation, int batchSize, int successCount, int failureCount, TimeSpan duration)
        {
            var message = $"Batch {operation}: Size={batchSize}, Success={successCount}, Failures={failureCount}, Duration={duration.TotalMilliseconds:F2}ms";
            
            // Batch operation logs only go to debug logger when debug logging is enabled
            if (_enableDebugLogging)
            {
                var logFilePath = GetDebugLogFilePath();
                Log("BATCH", message, logFilePath);
            }
        }

        private void Log(string logLevel, string message, string logFilePath)
        {
            var logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{logLevel}] {message}";
            try
            {
                if (logFilePath != null)
                {
                    // Ensure the directory exists
                    Directory.CreateDirectory(path: Path.GetDirectoryName(logFilePath));
                }

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
