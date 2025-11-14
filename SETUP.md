# EmailDeleter Setup Guide

This guide provides detailed step-by-step instructions for setting up and configuring the EmailDeleter application.

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Azure AD Setup](#azure-ad-setup)
3. [Application Configuration](#application-configuration)
4. [Testing Configuration](#testing-configuration)
5. [Production Deployment](#production-deployment)
6. [Troubleshooting](#troubleshooting)

## Prerequisites

### Software Requirements

1. **.NET 8.0 SDK**
   - Download from: https://dotnet.microsoft.com/download/dotnet/8.0
   - Verify installation: `dotnet --version`

2. **Visual Studio 2022** (Optional, for development)
   - Community, Professional, or Enterprise edition
   - Workload: .NET desktop development

3. **Microsoft Excel** (Optional, for viewing exports)
   - Or any Excel-compatible spreadsheet application

### Access Requirements

1. **Azure AD Admin Access**
   - Permission to create app registrations
   - Permission to grant admin consent for API permissions

2. **Microsoft 365 Access**
   - Access to mailboxes you want to process
   - Global Admin or Exchange Admin role (for shared mailboxes)

## Azure AD Setup

### Step 1: Create App Registration

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory**
3. Select **App registrations** from the left menu
4. Click **+ New registration**
5. Fill in the registration form:
   ```
   Name: EmailDeleter
   Supported account types: Accounts in this organizational directory only
   Redirect URI: Leave blank
   ```
6. Click **Register**

### Step 2: Note Application IDs

After registration, you'll see the Overview page. Note these values:

```
Application (client) ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Directory (tenant) ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

### Step 3: Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **+ New client secret**
3. Fill in:
   ```
   Description: EmailDeleter Secret
   Expires: Choose appropriate duration (e.g., 12 months)
   ```
4. Click **Add**
5. **IMPORTANT**: Copy the secret **Value** immediately (it won't be shown again)
   ```
   Secret Value: xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
   ```

### Step 4: Configure API Permissions

1. Go to **API permissions** in your app registration
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Select **Application permissions**
5. Add the following permissions:
   - **Mail.ReadWrite** (Read and write mail in all mailboxes)
   - **User.Read.All** (Read all users' full profiles)
6. Click **Add permissions**

### Step 5: Grant Admin Consent

⚠️ **This step requires Global Admin or Privileged Role Admin**

1. On the API permissions page, click **Grant admin consent for [Your Organization]**
2. Click **Yes** to confirm
3. Verify all permissions show green checkmarks in the "Status" column

### Step 6: Security Best Practices

1. **Restrict Access**: Add the app to a security group if needed
2. **Audit**: Enable Azure AD audit logging for this app
3. **Review**: Regularly review the app's activity in Azure AD logs
4. **Rotate Secrets**: Set a reminder to rotate the client secret before expiration

## Application Configuration

### Step 1: Clone/Download the Repository

```bash
git clone [repository-url]
cd EmailDeleter
```

### Step 2: Create graph-secrets.json

Create a new file `EmailDeleter/graph-secrets.json`:

```json
{
  "clientId": "your-application-client-id-from-step-2",
  "tenantId": "your-tenant-id-from-step-2",
  "secret": "your-client-secret-value-from-step-3"
}
```

**Example:**
```json
{
  "clientId": "a1b2c3d4-e5f6-g7h8-i9j0-k1l2m3n4o5p6",
  "tenantId": "z9y8x7w6-v5u4-t3s2-r1q0-p9o8n7m6l5k4",
  "secret": "AbC123~dEf456.GhI789-JkL012"
}
```

⚠️ **Security Note**: This file is already in `.gitignore`. Never commit it to version control.

### Step 3: Configure appsettings.json

Edit `EmailDeleter/appsettings.json`:

```json
{
  "CofigFile": {
    "path": "C:\\Path\\To\\Your\\DelMsgList.xlsx",
    "excelDir": "C:\\Path\\To\\Output\\Directory\\",
    "infoLogDir": "C:\\Path\\To\\Logs\\Directory\\"
  },
  "Logging": {
    "EnableDebugLogging": true
  }
}
```

**Path Guidelines:**
- Use absolute paths
- Use double backslashes (`\\`) on Windows
- Ensure directories exist or have permission to create them
- Example Windows path: `C:\\Projects\\EmailDeleter\\`
- Example Linux path: `/home/user/EmailDeleter/`

### Step 4: Create Email Configuration Excel File

Create an Excel file at the path specified in `appsettings.json`:

**File Name**: `DelMsgList.xlsx`

**Sheet Structure**: First worksheet with the following columns:

| Column # | Header | Description | Values |
|----------|--------|-------------|--------|
| A | Email | Email address | user@domain.com |
| B | Inbox | Days threshold for Inbox | Number (e.g., 30) |
| C | Deleted | Days threshold for Deleted Items | Number (e.g., 30) |
| D | Sent | Days threshold for Sent Items | Number (e.g., 30) |
| E | Body | Include email body in export | 0 or 1 |
| F | Attachment | Only process emails with attachments | 0 or 1 |
| G | Read | Only process read emails | 0 or 1 |

**Example Configuration:**

| Email | Inbox | Deleted | Sent | Body | Attachment | Read |
|-------|-------|---------|------|------|------------|------|
| john.doe@company.com | 30 | 30 | 30 | 1 | 0 | 0 |
| jane.smith@company.com | 60 | 30 | 60 | 1 | 1 | 0 |
| archive@company.com | 90 | 30 | 90 | 0 | 0 | 1 |

**Configuration Explanation:**
- Row 2: Delete all emails older than 30 days from all folders
- Row 3: Delete emails with attachments older than 60 days (Inbox/Sent)
- Row 4: Delete only read emails older than 90 days

### Step 5: Create Required Directories

Ensure the directories specified in `appsettings.json` exist:

**Windows:**
```cmd
mkdir C:\Path\To\Output\Directory
mkdir C:\Path\To\Logs\Directory
```

**Linux/Mac:**
```bash
mkdir -p /path/to/output/directory
mkdir -p /path/to/logs/directory
```

### Step 6: Build the Application

```bash
cd EmailDeleter
dotnet restore
dotnet build --configuration Release
```

Verify successful build:
```
Build succeeded.
    0 Warning(s)
    0 Error(s)
```

## Testing Configuration

### Test 1: Configuration Loading

1. Create a test configuration with a single email
2. Set very high day thresholds (e.g., 365 days) to avoid accidental deletion
3. Run the application in debug mode:

```bash
dotnet run --configuration Debug
```

4. Check logs for:
   - ✅ "Configuration loaded successfully"
   - ✅ "Starting processing for email: [your-email]"
   - ✅ No authentication errors

### Test 2: Authentication

Review the log files:

```
InfoLog-YYYY-MM-DD.log
```

Look for successful authentication:
- ✅ No "Service exception" errors
- ✅ Messages about email fetching starting
- ✅ Count of emails processed (even if 0)

### Test 3: Dry Run (Optional)

To test without actually deleting emails, you can comment out the deletion code in `Program.cs`:

**Lines to comment (183-186 and 188-189):**
```csharp
// Temporarily disable deletion for testing
// await graphClient.Users[config.email]
//     .Messages[message.Id]
//     .DeleteAsync(Microsoft.Graph.Models.DeletionMode.HardDelete);
```

### Test 4: Verify Excel Export

1. After a test run, check the output directory
2. Open the generated Excel file (e.g., `john.xlsx`)
3. Verify:
   - ✅ Headers are present
   - ✅ Email data is populated
   - ✅ Dates are formatted correctly

## Production Deployment

### Pre-Deployment Checklist

- [ ] Azure AD app configured with correct permissions
- [ ] Admin consent granted
- [ ] `graph-secrets.json` created with valid credentials
- [ ] `appsettings.json` configured with production paths
- [ ] Email configuration Excel file created and validated
- [ ] Output and log directories created with appropriate permissions
- [ ] Application tested with debug logging enabled
- [ ] Backup of mailbox data (if critical)

### Deployment Steps

1. **Disable Debug Logging** (for performance)
   ```json
   {
     "Logging": {
       "EnableDebugLogging": false
     }
   }
   ```

2. **Build Release Version**
   ```bash
   dotnet build --configuration Release
   dotnet publish --configuration Release --output ./publish
   ```

3. **Deploy Files**
   - Copy contents of `./publish` to production server
   - Copy `graph-secrets.json` securely
   - Copy `appsettings.json` with production paths
   - Copy email configuration Excel file

4. **Set Up Scheduled Execution** (Optional)

   **Windows Task Scheduler:**
   - Create new task
   - Trigger: Daily at desired time
   - Action: Start program `dotnet.exe`
   - Arguments: `path\to\EmailDeleter.dll`

   **Linux Cron:**
   ```bash
   crontab -e
   # Run daily at 2 AM
   0 2 * * * /usr/bin/dotnet /path/to/EmailDeleter.dll >> /path/to/cron.log 2>&1
   ```

5. **Monitor Initial Runs**
   - Check log files after each run
   - Verify Excel exports
   - Review error logs
   - Monitor performance metrics

### Security Hardening

1. **File Permissions**
   ```bash
   # Linux
   chmod 600 graph-secrets.json
   chmod 644 appsettings.json
   chmod 700 logs/
   ```

2. **Encrypt Secrets** (Advanced)
   - Consider using Azure Key Vault
   - Or Windows DPAPI for local encryption

3. **Network Security**
   - Ensure firewall allows HTTPS to `graph.microsoft.com`
   - Use corporate proxy if required

4. **Audit Trail**
   - Enable Azure AD audit logging
   - Review logs regularly
   - Set up alerts for suspicious activity

## Troubleshooting

### Issue: "Configuration file not found"

**Solution:**
1. Verify path in `appsettings.json` is correct
2. Check file permissions
3. Use absolute path, not relative
4. Ensure double backslashes on Windows

### Issue: "Authentication failed" or "401 Unauthorized"

**Solution:**
1. Verify credentials in `graph-secrets.json`
2. Check Azure AD app permissions
3. Ensure admin consent is granted
4. Verify tenant ID is correct
5. Try regenerating client secret

### Issue: "403 Forbidden"

**Solution:**
1. Verify API permissions are granted
2. Check if admin consent is provided
3. Ensure service account has mailbox access
4. Review Azure AD conditional access policies

### Issue: "No emails deleted" or "0 emails processed"

**Solution:**
1. Check date thresholds in configuration
2. Verify filter criteria (attachment, read status)
3. Enable debug logging to see filter details
4. Manually verify emails exist matching criteria

### Issue: "Excel file locked" error

**Solution:**
1. Close Excel file if open
2. Application will create timestamped backup file
3. Check file permissions
4. Ensure output directory is writable

### Issue: Performance degradation with large mailboxes

**Solution:**
1. Reduce batch size in `Program.cs` (line 287 and 444)
2. Process fewer accounts per run
3. Adjust timeout values if needed
4. Enable debug logging to identify bottlenecks

### Issue: "Graph API rate limiting" errors

**Solution:**
1. Reduce batch size
2. Add delays between operations
3. Process accounts sequentially, not in parallel
4. Contact Microsoft support for rate limit increases

## Support and Resources

### Documentation
- [Microsoft Graph API Documentation](https://docs.microsoft.com/graph/)
- [Azure AD App Registration Guide](https://docs.microsoft.com/azure/active-directory/develop/quickstart-register-app)
- [.NET 8.0 Documentation](https://docs.microsoft.com/dotnet/)

### Logging
- See `LOGGING_IMPROVEMENTS.md` for detailed logging information
- Review log files in the configured log directory
- Enable debug logging for troubleshooting

### Getting Help
1. Check log files for detailed error messages
2. Review this setup guide
3. Consult Microsoft Graph API documentation
4. Create an issue in the repository (if applicable)

## Next Steps

After successful setup:
1. Run a test with a single mailbox
2. Verify Excel exports
3. Review logs
4. Gradually increase scope
5. Set up scheduled execution
6. Implement monitoring and alerting
