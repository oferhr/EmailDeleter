# EmailDeleter

A .NET application for automated bulk email deletion from Microsoft 365/Exchange Online mailboxes using Microsoft Graph API.

## Overview

EmailDeleter is a command-line tool designed to help organizations and individuals manage email retention by automatically deleting emails based on configurable criteria such as age, folder location, and attachments. The application processes emails from multiple mailboxes, moves them to Deleted Items, performs hard deletion, and exports metadata to Excel for audit purposes.

## Features

- **Bulk Email Processing**: Process multiple email accounts from a single configuration file
- **Flexible Filtering**: Delete emails based on:
  - Age (days old)
  - Folder location (Inbox, Sent Items, Deleted Items)
  - Attachment presence
- **Batch Operations**: Efficient batch processing with Microsoft Graph API
- **Audit Trail**: Export deleted email metadata to Excel files
- **Comprehensive Logging**: Multiple log levels (INFO, DEBUG, WARNING, ERROR, PERFORMANCE, BATCH)
- **Error Handling**: Robust error handling with automatic retry for batch operations
- **Performance Tracking**: Built-in performance monitoring for all operations

## Prerequisites

- .NET 8.0 SDK or later
- Microsoft Azure AD application registration with Microsoft Graph API permissions
- Microsoft 365/Exchange Online environment

### Required Microsoft Graph API Permissions

The Azure AD application requires the following delegated or application permissions:
- `Mail.ReadWrite` - To read and delete emails
- `Mail.ReadWrite.Shared` - If processing shared mailboxes
- `User.Read.All` - To access user mailboxes

## Project Structure

```
EmailDeleter/
├── EmailDeleter.sln              # Visual Studio solution file
├── EmailDeleter/
│   ├── Program.cs                # Main application logic
│   ├── logger.cs                 # Logging implementation
│   ├── EmailDeleter.csproj       # Project configuration
│   ├── appsettings.json          # Application settings
│   ├── graph-secrets.json        # Azure AD credentials (not in repo)
│   └── LOGGING_IMPROVEMENTS.md   # Logging documentation
└── README.md                     # This file
```

## Setup

### 1. Azure AD App Registration

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to Azure Active Directory > App registrations > New registration
3. Configure the application:
   - Name: EmailDeleter
   - Supported account types: Choose appropriate option
   - Redirect URI: Not required for this application
4. Note the **Application (client) ID** and **Directory (tenant) ID**
5. Create a client secret:
   - Go to Certificates & secrets > New client secret
   - Save the secret value securely
6. Grant API permissions:
   - Go to API permissions > Add a permission
   - Select Microsoft Graph > Application permissions
   - Add: `Mail.ReadWrite`, `User.Read.All`
   - Click "Grant admin consent"

### 2. Configuration Files

#### Create `graph-secrets.json`
Create this file in the EmailDeleter directory (it's in .gitignore):

```json
{
  "clientId": "your-application-client-id",
  "tenantId": "your-tenant-id",
  "secret": "your-client-secret-value"
}
```

#### Configure `appsettings.json`
Update paths according to your environment:

```json
{
  "CofigFile": {
    "path": "C:\\Path\\To\\DelMsgList.xlsx",
    "excelDir": "C:\\Path\\To\\Output\\",
    "infoLogDir": "C:\\Path\\To\\Logs\\"
  },
  "Logging": {
    "EnableDebugLogging": true
  }
}
```

#### Create Email Configuration Excel File

Create an Excel file (e.g., `DelMsgList.xlsx`) with the following columns:

| Column | Description | Type | Example |
|--------|-------------|------|---------|
| Email | Email address to process | String | user@domain.com |
| Inbox | Days threshold for Inbox | Integer | 30 |
| Deleted | Days threshold for Deleted Items | Integer | 30 |
| Sent | Days threshold for Sent Items | Integer | 30 |
| Body | Include body in export (1=yes, 0=no) | Integer | 1 |
| Attachment | Only delete emails with attachments (1=yes, 0=no) | Integer | 0 |
| Read | Only delete read emails (1=yes, 0=no) | Integer | 0 |

**Example:**
| Email | Inbox | Deleted | Sent | Body | Attachment | Read |
|-------|-------|---------|------|------|------------|------|
| user1@domain.com | 30 | 30 | 30 | 1 | 0 | 0 |
| user2@domain.com | 60 | 30 | 60 | 1 | 1 | 0 |

### 3. Build and Run

```bash
cd EmailDeleter
dotnet restore
dotnet build
dotnet run
```

Or use Visual Studio:
1. Open `EmailDeleter.sln`
2. Build the solution (F6)
3. Run the application (F5)

## How It Works

1. **Initialization**: Loads configuration and initializes logging
2. **Configuration Reading**: Reads email accounts and processing rules from Excel
3. **Email Fetching**: For each account and folder:
   - Queries Microsoft Graph API with filters
   - Retrieves emails matching criteria (age, attachments)
   - Processes in batches of 10 messages per page
4. **Email Moving**: Moves emails from Inbox/Sent Items to Deleted Items
5. **Hard Delete**: Permanently deletes emails from Deleted Items folder
6. **Export**: Saves email metadata (from, to, subject, body, date) to Excel
7. **Logging**: Records all operations with performance metrics

## Architecture

### Main Components

- **Program.cs**: Main application logic
  - `Main()`: Entry point, orchestrates the deletion process
  - `fetchEmails()`: Retrieves emails matching criteria
  - `moveToDeleted()`: Batch moves emails to Deleted Items
  - `ProcessDeleteBatchAsync()`: Batch deletes emails permanently
  - `WriteToExcel()`: Exports email metadata to Excel
  - `ReadConfig()`: Loads configuration from Excel

- **logger.cs**: Logging infrastructure
  - `SimpleLogger`: Custom logger with multiple log levels
  - Separate log files for different log types
  - Configurable debug logging

### Data Models

- **ConfigData**: Email account configuration
- **EmailData**: Email metadata for processing and export

## Logging

The application provides comprehensive logging at multiple levels. See [LOGGING_IMPROVEMENTS.md](EmailDeleter/LOGGING_IMPROVEMENTS.md) for detailed information.

### Log Files

- `log-YYYY-MM-DD.log` - Error logs
- `InfoLog-YYYY-MM-DD.log` - Informational logs
- `DebugLog-YYYY-MM-DD.log` - Detailed debug logs (when enabled)

### Enable/Disable Debug Logging

Set in `appsettings.json`:
```json
{
  "Logging": {
    "EnableDebugLogging": false  // Set to true for detailed logs
  }
}
```

## Security Considerations

⚠️ **Important Security Notes:**

1. **Credentials**: Never commit `graph-secrets.json` to version control
2. **Permissions**: Use least-privilege principle for Graph API permissions
3. **Audit**: Review deleted emails in Excel exports before permanent deletion
4. **Access Control**: Restrict access to configuration files and logs
5. **Service Account**: Consider using a dedicated service account for the Azure AD app
6. **Log Security**: Protect log files as they may contain email subjects and metadata

## Error Handling

The application includes robust error handling:
- Automatic retry for failed batch operations
- Fallback to individual operations if batch fails
- Detailed error logging with stack traces
- Graceful degradation for network issues

## Performance

- **Batch Processing**: Processes up to 5 emails per batch request
- **Pagination**: Handles large mailboxes with automatic pagination
- **Performance Logging**: Tracks duration of all major operations
- **Memory Management**: Proper disposal of resources and Excel workbooks

## Troubleshooting

### Common Issues

1. **Authentication Failed**
   - Verify `graph-secrets.json` credentials
   - Check Azure AD app permissions
   - Ensure admin consent is granted

2. **No Emails Deleted**
   - Check date threshold in configuration
   - Verify filter criteria (attachments, read status)
   - Review logs for filter details

3. **Excel File Locked**
   - Close Excel file if open
   - Application creates timestamped backup if file is locked

4. **Performance Issues**
   - Reduce batch size in code if timeouts occur
   - Enable debug logging to identify bottlenecks
   - Check network connectivity to Microsoft Graph API

## Version History

- **1.0.1.5** (Current)
  - Enhanced logging with multiple log levels
  - Performance tracking and monitoring
  - Improved error handling
  - Batch operation statistics

- **1.0.0**
  - Initial release
  - Basic email deletion functionality
  - Excel export capabilities

## Contributing

When contributing to this project:
1. Follow existing code style and patterns
2. Add appropriate logging for new features
3. Update documentation for significant changes
4. Test with various email scenarios
5. Never commit secrets or credentials

## License

[Specify your license here]

## Support

For issues, questions, or contributions, please [create an issue](link-to-issues) in the repository.
