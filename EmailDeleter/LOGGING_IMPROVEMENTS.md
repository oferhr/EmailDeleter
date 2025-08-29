# EmailDeleter Logging Improvements

## Overview
The EmailDeleter application has been enhanced with comprehensive logging capabilities to provide better visibility into operations, debugging capabilities, and performance monitoring.

## Logging Levels

### 1. INFO Level
- Application startup and shutdown
- Configuration loading status
- Email processing progress
- Batch operation summaries
- Performance metrics

### 2. DEBUG Level
- Detailed operation steps
- Individual message processing
- API request details
- Configuration parsing
- File operations

### 3. WARNING Level
- Non-critical issues
- Configuration fallbacks
- File access issues
- API response warnings

### 4. ERROR Level
- Exceptions and stack traces
- Service failures
- Configuration errors
- File system errors

### 5. PERFORMANCE Level
- Operation timing
- Batch processing duration
- API call performance
- Memory usage tracking

### 6. BATCH Level
- Batch operation statistics
- Success/failure counts
- Batch processing duration

## Enhanced Logging Features

### 1. Performance Tracking
- **Operation Timing**: All major operations now include timing information
- **Batch Statistics**: Detailed success/failure counts for batch operations
- **API Call Monitoring**: Duration tracking for Graph API calls

### 2. Debug Information
- **Message Processing**: Individual email processing details
- **Pagination Progress**: Page-by-page processing information
- **Configuration Details**: Detailed configuration parsing logs
- **File Operations**: Excel file read/write operations

### 3. Error Context
- **Service Exceptions**: Detailed Microsoft Graph API error information
- **Request/Response Details**: API call details for debugging
- **Configuration Validation**: Configuration file validation logs

### 4. Operational Logs
- **Startup/Shutdown**: Application lifecycle logging
- **Configuration Loading**: Configuration file processing
- **Memory Management**: Resource cleanup logging

## Configuration

### appsettings.json
```json
{
  "CofigFile": {
    "path": "C:\\Intel\\DelMsgList.xlsx",
    "excelDir": "C:\\Intel\\",
    "infoLogDir": "C:\\Intel\\"
  },
  "Logging": {
    "EnableDebugLogging": true,
    "LogLevel": {
      "Default": "Information",
      "Microsoft": "Warning",
      "Microsoft.Hosting.Lifetime": "Information"
    }
  }
}
```

### Log Files
- **Error Logs**: `log-YYYY-MM-DD.log` - Contains all ERROR level messages
- **Info Logs**: `InfoLog-YYYY-MM-DD.log` - Contains INFO, WARNING, PERFORMANCE, and BATCH messages
- **Debug Logs**: `DebugLog-YYYY-MM-DD.log` - Contains DEBUG level messages (when enabled)

## Key Improvements Made

### 1. Enhanced SimpleLogger Class
- Added `LogWarning()`, `LogDebug()`, `LogPerformance()`, and `LogBatchOperation()` methods
- Improved timestamp precision (includes milliseconds)
- Configurable debug logging enable/disable
- Separate log files for different log levels

### 2. Application Lifecycle Logging
- Application startup and shutdown logging
- Configuration validation and loading
- Error handling with detailed context

### 3. Email Processing Logging
- Individual message processing details
- Pagination progress tracking
- Batch operation statistics
- Performance metrics for each operation

### 4. API Operation Logging
- Graph API call details
- Request/response information
- Error context for failed operations
- Rate limiting and retry information

### 5. File Operation Logging
- Excel file read/write operations
- Configuration file processing
- File lock handling
- Directory creation and validation

## Usage Examples

### Performance Monitoring
```
2024-01-15 10:30:45.123 [PERFORMANCE] Performance: Email fetch for user@domain.com in Inbox took 1250.45ms Total emails processed: 45
2024-01-15 10:30:46.456 [PERFORMANCE] Performance: Batch move for 45 messages took 890.12ms
2024-01-15 10:30:47.789 [PERFORMANCE] Performance: Delete batch for 45 messages took 567.34ms
```

### Batch Operations
```
2024-01-15 10:30:46.456 [BATCH] Batch Move to DeletedItems: Size=45, Success=43, Failures=2, Duration=890.12ms
2024-01-15 10:30:47.789 [BATCH] Batch Delete from DeletedItems: Size=43, Success=42, Failures=1, Duration=567.34ms
```

### Debug Information
```
2024-01-15 10:30:45.123 [DEBUG] Processing page 1 with 10 messages
2024-01-15 10:30:45.124 [DEBUG] Processing message ID: AAMkAGI2..., Subject: Meeting reminder...
2024-01-15 10:30:45.125 [DEBUG] Successfully moved message AAMkAGI2...
```

## Benefits

1. **Troubleshooting**: Detailed logs help identify issues quickly
2. **Performance Monitoring**: Track operation performance and identify bottlenecks
3. **Audit Trail**: Complete record of all operations performed
4. **Debugging**: Detailed debug information for development and testing
5. **Monitoring**: Real-time visibility into application health and performance

## Recommendations

1. **Production Use**: Set `EnableDebugLogging` to `false` in production to reduce log volume
2. **Log Rotation**: Implement log rotation to manage disk space
3. **Monitoring**: Set up alerts for ERROR level messages
4. **Performance Analysis**: Use PERFORMANCE logs to identify optimization opportunities
5. **Security**: Ensure log files are properly secured and not accessible to unauthorized users
