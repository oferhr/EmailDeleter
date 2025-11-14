namespace EmailDeleter
{
    /// <summary>
    /// Simple file-based logger with multiple log levels
    /// Supports separate log files for different log types (error, info, debug)
    /// Thread-safe file writing with automatic daily log rotation
    /// </summary>
    public class SimpleLogger
    {
        private readonly string _baseFileName;
        private readonly string _infoLogPath;
        private readonly bool _enableDebugLogging;

        /// <summary>
        /// Initializes a new instance of the SimpleLogger
        /// </summary>
        /// <param name="baseFileName">Base filename for log files</param>
        /// <param name="infoLogParg">Directory path for info and debug logs</param>
        /// <param name="enableDebugLogging">Enable detailed debug logging (default: false)</param>
        public SimpleLogger(string baseFileName, string infoLogParg, bool enableDebugLogging = false)
        {
            _baseFileName = baseFileName;
            _infoLogPath = infoLogParg;
            _enableDebugLogging = enableDebugLogging;
        }

        /// <summary>
        /// Gets the log file path for error logs with current date
        /// Format: log-YYYY-MM-DD.log
        /// </summary>
        private string GetLogFilePath()
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            return $"log-{date}.log";
        }

        /// <summary>
        /// Gets the log file path for info logs with current date
        /// Format: InfoLog-YYYY-MM-DD.log
        /// </summary>
        private string GetInfoLogFilePath()
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            return Path.Combine(_infoLogPath, $"InfoLog-{date}.log");
        }

        /// <summary>
        /// Gets the log file path for debug logs with current date
        /// Format: DebugLog-YYYY-MM-DD.log
        /// </summary>
        private string GetDebugLogFilePath()
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            return Path.Combine(_infoLogPath, $"DebugLog-{date}.log");
        }

        /// <summary>
        /// Logs an informational message
        /// Writes to info log file (and debug log if debug logging is enabled)
        /// </summary>
        /// <param name="message">Message to log</param>
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

        /// <summary>
        /// Logs a warning message (non-critical issues)
        /// Writes to info log file (and debug log if debug logging is enabled)
        /// </summary>
        /// <param name="message">Warning message to log</param>
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

        /// <summary>
        /// Logs a debug message (detailed diagnostic information)
        /// Only writes to log file if debug logging is enabled
        /// </summary>
        /// <param name="message">Debug message to log</param>
        public void LogDebug(string message)
        {
            if (_enableDebugLogging)
            {
                var logFilePath = GetDebugLogFilePath();
                Log("DEBUG", message, logFilePath);
            }
        }

        /// <summary>
        /// Logs an error message with optional exception details
        /// Writes to error log file (and debug log if debug logging is enabled)
        /// Includes exception message and stack trace if provided
        /// </summary>
        /// <param name="message">Error description</param>
        /// <param name="ex">Optional exception object</param>
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

        /// <summary>
        /// Logs performance metrics for operations
        /// Only writes to log if debug logging is enabled
        /// </summary>
        /// <param name="operation">Name of the operation being measured</param>
        /// <param name="duration">Time taken to complete the operation</param>
        /// <param name="additionalInfo">Optional additional context information</param>
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

        /// <summary>
        /// Logs statistics for batch operations
        /// Only writes to log if debug logging is enabled
        /// </summary>
        /// <param name="operation">Name of the batch operation</param>
        /// <param name="batchSize">Total number of items in the batch</param>
        /// <param name="successCount">Number of successful operations</param>
        /// <param name="failureCount">Number of failed operations</param>
        /// <param name="duration">Total time taken for the batch</param>
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

        /// <summary>
        /// Core logging method that writes formatted messages to a file
        /// Thread-safe file writing with automatic directory creation
        /// </summary>
        /// <param name="logLevel">Log level (INFO, WARNING, ERROR, DEBUG, PERFORMANCE, BATCH)</param>
        /// <param name="message">Message to log</param>
        /// <param name="logFilePath">Full path to the log file</param>
        private void Log(string logLevel, string message, string logFilePath)
        {
            // Format: YYYY-MM-DD HH:mm:ss.fff [LEVEL] Message
            var logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{logLevel}] {message}";
            try
            {
                if (logFilePath != null)
                {
                    // Ensure the directory exists before writing
                    Directory.CreateDirectory(path: Path.GetDirectoryName(logFilePath));
                }

                // Append the log message to the file (creates file if doesn't exist)
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                // Fallback to console if file writing fails
                Console.WriteLine($"Failed to write to log file: {ex.Message}");
            }
        }
    }
}
