using Azure.Identity;
using ClosedXML.Excel;
using EmailDeleter;
using HtmlAgilityPack;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System.Text.Json;

/// <summary>
/// EmailDeleter - Automated bulk email deletion tool for Microsoft 365/Exchange Online
/// Uses Microsoft Graph API to process and delete emails based on configurable criteria
/// </summary>
class Program
{
    // Collection of emails currently being processed
    private static List<EmailData> _emails = new List<EmailData>();

    // Logger instances for different log types
    private static SimpleLogger logger = null!;
    private static SimpleLogger infoLogger = null!;

    // Track deletion counts per folder
    private static Dictionary<string, int> counts = new Dictionary<string, int>();

    // Configuration constants
    private const int BATCH_SIZE = 5;
    private const int PAGE_SIZE = 10;
    private const int DEFAULT_DAYS_THRESHOLD = 30;
    private const int DEFAULT_ATTACHMENT_FLAG = 1;
    private const int DEFAULT_BODY_FLAG = 1;

    /// <summary>
    /// Application entry point
    /// Processes multiple email accounts and deletes emails based on configuration
    /// </summary>
    /// <param name="args">Command line arguments (not currently used)</param>
    static async Task Main(string[] args)
    {
        var startTime = DateTime.UtcNow;
        
        // Initialize logger first
        InitializeLogger();
        logger.LogInfo("EmailDeleter application started");
        
        try
        {
            var ConfigData = ReadConfig();
            logger.LogInfo($"Configuration loaded successfully. Processing {ConfigData.Count} email accounts");
            
            foreach (var config in ConfigData)
            {
                logger.LogInfo($"Starting processing for email: {config.email}");
                var configStartTime = DateTime.UtcNow;
                
                counts["Inbox"] = await fetchEmails(config, "Inbox", config.inbox, false);
                counts["SentItems"] = await fetchEmails(config, "SentItems", config.sent, false);
                counts["DeletedItems"] = await fetchEmails(config, "DeletedItems", config.deleted, true);
                
                var configDuration = DateTime.UtcNow - configStartTime;
                logger.LogPerformance($"Processing {config.email}", configDuration, $"Inbox: {counts["Inbox"]}, SentItems: {counts["SentItems"]}, DeletedItems: {counts["DeletedItems"]}");
                infoLogger.LogInfo($"Emails deleted: Inbox: {counts["Inbox"]}, SentItems: {counts["SentItems"]}, DeletedItems: {counts["DeletedItems"]}");
                var totalDuration = DateTime.UtcNow - startTime;
                logger.LogPerformance("Total application execution for " + config.email, totalDuration, $"Total emails processed: {counts.Values.Sum()}");
            }
            
            
            
            logger.LogInfo("EmailDeleter application completed successfully");
        }
        catch (Exception ex)
        {
            logger.LogError("Application failed with unexpected error", ex);
            throw;
        }
    }
    
    /// <summary>
    /// Initializes the logging system with configuration from appsettings.json
    /// Falls back to basic logging if configuration fails
    /// </summary>
    private static void InitializeLogger()
    {
        try
        {
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();

            // Get log directory from configuration
            var path = config["ConfigFile:infoLogDir"];
            if (string.IsNullOrEmpty(path))
            {
                logger.LogWarning("Info log directory not configured, using default path");
                path = AppContext.BaseDirectory;
            }

            // Check if debug logging is enabled
            var enableDebugLogging = config["Logging:EnableDebugLogging"] != null
                ? bool.Parse(config["Logging:EnableDebugLogging"])
                : false;

            // Initialize logger instances
            logger = new SimpleLogger("log", path, enableDebugLogging);
            infoLogger = new SimpleLogger("log", path, enableDebugLogging);
            logger.LogInfo($"Info logger initialized with path: {path}");
        }
        catch (Exception ex)
        {
            // Fallback to basic logger if configuration fails
            logger = new SimpleLogger("log", string.Empty, false);
            Console.WriteLine($"Failed to initialize logger with configuration, using fallback: {ex.Message}");
        }
    }
    
    /// <summary>
    /// Fetches emails from a specific folder based on age and criteria
    /// Processes emails in batches and exports metadata to Excel
    /// </summary>
    /// <param name="config">Email account configuration</param>
    /// <param name="dir">Folder name (Inbox, SentItems, DeletedItems)</param>
    /// <param name="days">Number of days threshold (emails older than this will be processed)</param>
    /// <param name="isDeleted">True if processing DeletedItems folder (skips move operation)</param>
    /// <returns>Number of emails processed</returns>
    static async Task<int> fetchEmails(ConfigData config, string dir, int days, bool isDeleted)
    {
        var startTime = DateTime.UtcNow;
        logger.LogInfo($"Starting email fetch for {config.email} in {dir} folder (days: {days}, isDeleted: {isDeleted})");

        // Load Azure AD credentials from graph-secrets.json
        // This file is kept separate to avoid committing secrets to version control
        var secretsConfig = new ConfigurationBuilder()
           .SetBasePath(AppContext.BaseDirectory)
           .AddJsonFile("graph-secrets.json", optional: false, reloadOnChange: true)
           .Build();
        var clientId = secretsConfig["clientId"];
        var tenantId = secretsConfig["tenantId"];
        var secret = secretsConfig["secret"];

        logger.LogDebug($"Authenticating with Azure AD for user: {config.email}");
        infoLogger.LogInfo($"Fetching emails for {config.email} in {dir} folder.");

        // Authenticate with Azure AD using client credentials flow
        var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, secret);
        var graphClient = new GraphServiceClient(clientSecretCredential);
        int counter = 0;
        try
        {
            // Build filter for Microsoft Graph API query
            // Calculate date threshold based on days parameter
            var dateThreshold = DateTime.Now.AddDays(days*-1).ToString("yyyy-MM-ddTHH:mm:ssZ");

            // Add attachment filter if configured
            var attachment = config.attachment ? "hasAttachments eq true and" : "";
            var filter = $"{attachment} receivedDateTime lt {dateThreshold}";

            logger.LogDebug($"Using filter: {filter}");
            logger.LogDebug($"Date threshold: {dateThreshold}");

            // Query Microsoft Graph API for messages matching criteria
            // Process in batches of 10 messages per page
            var page = await graphClient.Users[config.email]
                .MailFolders[dir]
                .Messages
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                    requestConfiguration.QueryParameters.Select = new string[] { "subject,body,receivedDateTime,from,toRecipients,isRead" };
                    requestConfiguration.QueryParameters.Top = PAGE_SIZE;
                });


            // Process all pages of results
            var pageCount = 0;
            while (page != null)
            {
                pageCount++;
                logger.LogDebug($"Processing page {pageCount} with {page.Value?.Count ?? 0} messages");
                
                foreach (var message in page.Value)
                {
                    if (message?.Id == null)
                    {
                        logger.LogWarning("Skipping message with null ID");
                        continue;
                    }
                    
                    var subjectPreview = message.Subject?.Length > 50 
                        ? message.Subject.Substring(0, 50) + "..." 
                        : message.Subject ?? "No Subject";
                    logger.LogDebug($"Processing message ID: {message.Id}, Subject: {subjectPreview}");
                    
                    var htmlBody = message.Body?.Content ?? "No content available.";
                    var plainTextBody = ExtractPlainTextFromHtml(htmlBody);
                    var emailData = new EmailData
                    {
                        from = message.From?.EmailAddress?.Address ?? "Unknown",
                        toRecipients = message.ToRecipients?.Select(r => r.EmailAddress?.Address).Where(addr => !string.IsNullOrEmpty(addr)).ToList() ?? new List<string>(),
                        subject = message.Subject ?? "No Subject",
                        receivedDateTime = message.ReceivedDateTime?.ToString("yyyy-MM-ddTHH:mm:ssZ") ?? "",
                        body = plainTextBody.Trim(),
                        id = message.Id,
                        source = dir // Set the source folder
                    };

                    _emails.Add(emailData);
                }

                // Process the batch of emails
                var isOK = true;
                var newEmails = new List<EmailData>();

                // Step 1: Move emails to DeletedItems (unless already in DeletedItems)
                if (!isDeleted)
                {
                    try
                    {
                        logger.LogDebug($"Moving {_emails.Count} messages to DeletedItems folder");
                        var moveStartTime = DateTime.UtcNow;
                        newEmails = await moveToDeleted(_emails, graphClient, config.email);
                        var moveDuration = DateTime.UtcNow - moveStartTime;
                        logger.LogPerformance($"Move to DeletedItems for {_emails.Count} messages", moveDuration);
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Error moving messages to DeletedItems: {ex.Message}", ex);
                        isOK = false;
                    }
                }
                else
                {
                    // Already in DeletedItems, no need to move
                    logger.LogDebug($"Skipping move operation for DeletedItems folder");
                    newEmails = _emails;
                }

                // Step 2: Permanently delete emails (hard delete)
                if (isOK)
                {
                    logger.LogDebug($"Processing delete batch for {newEmails.Count} messages");
                    var deleteStartTime = DateTime.UtcNow;
                    await ProcessDeleteBatchAsync(newEmails, graphClient, config.email);
                    var deleteDuration = DateTime.UtcNow - deleteStartTime;
                    logger.LogPerformance($"Delete batch for {newEmails.Count} messages", deleteDuration);
                }

                // Step 3: Export email metadata to Excel for audit trail
                var excelStartTime = DateTime.UtcNow;
                var excelSuccess = WriteToExcel(config.email);
                var excelDuration = DateTime.UtcNow - excelStartTime;
                logger.LogPerformance($"Excel write for {_emails.Count} emails", excelDuration, $"Success: {excelSuccess}");
                
                infoLogger.LogInfo($"{_emails.Count.ToString()} Emails deleted for {config.email} in {dir} folder.");
                counter += _emails.Count;
                _emails = new List<EmailData>();
                // Fetch the next page using the OdataNextLink
                if (!string.IsNullOrEmpty(page.OdataNextLink))
                {
                    logger.LogDebug($"Fetching next page: {page.OdataNextLink}");
                    var nextPageRequest = new RequestInformation
                    {
                        HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
                        UrlTemplate = page.OdataNextLink
                    };

                    page = await graphClient.RequestAdapter.SendAsync<MessageCollectionResponse>(
                        nextPageRequest,
                        MessageCollectionResponse.CreateFromDiscriminatorValue, // Factory method
                        default, // No additional parsable factories
                        CancellationToken.None // Cancellation token
                    );
                }
                else
                {
                    logger.LogDebug("No more pages to fetch");
                    page = null; // No more pages
                }
            }
            logger.LogDebug("finish page with page_count of: " + pageCount);
        }
        catch (ServiceException ex)
        {
            logger.LogError($"Service exception while fetching messages: {ex.Message}", ex);
            logger.LogDebug($"Service exception details - Status: {ex.ResponseStatusCode},  Message: {ex.Message}");
        }
        catch (Exception ex)
        {
            logger.LogError($"Unexpected error while fetching messages: {ex.Message}", ex);
        }
        finally
        {
            var totalDuration = DateTime.UtcNow - startTime;
            logger.LogPerformance($"Email fetch for {config.email} in {dir}", totalDuration, $"Total emails processed: {counter}");
        }
        
        return counter;
    }

    /// <summary>
    /// Moves a batch of emails to the DeletedItems folder using Microsoft Graph API batch operations
    /// Updates email objects with new IDs after successful move
    /// Falls back to individual moves if batch operation fails
    /// </summary>
    /// <param name="emails">List of emails to move</param>
    /// <param name="graphClient">Authenticated Graph API client</param>
    /// <param name="email">Email address of the mailbox being processed</param>
    /// <returns>Updated list of emails with new IDs</returns>
    static async Task<List<EmailData>> moveToDeleted(List<EmailData> emails, GraphServiceClient graphClient, string email)
    {
        logger.LogDebug($"Starting batch move operation for {emails.Count} emails to DeletedItems");

        var successCount = 0;
        var failureCount = 0;
        var batchStartTime = DateTime.UtcNow;

        // Process emails in smaller batches to avoid timeout issues
        var batches = emails.Chunk(BATCH_SIZE);
        
        foreach (var batch in batches)
        {
            try
            {
                logger.LogDebug($"Processing batch of {batch.Count()} emails");
                
                var batchRequestContent = new BatchRequestContentCollection(graphClient);
                var requestDictionary = new Dictionary<string, string>();
                
                foreach (var message in batch)
                {
                    logger.LogDebug($"Preparing move request for message ID: {message.id}");
                    
                    // Create the move request using the proper Graph SDK method
                    var moveRequest = graphClient.Users[email]
                        .Messages[message.id]
                        .Move
                        .ToPostRequestInformation(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                        {
                            DestinationId = "DeletedItems"
                        });

                    // Convert RequestInformation to HttpRequestMessage
                    var moveHttpRequest = await graphClient.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
                        moveRequest,
                        default(CancellationToken)
                    );

                    // Add the request to the batch using the proper method
                    var requestId = Guid.NewGuid().ToString();
                    batchRequestContent.AddBatchRequestStep(new BatchRequestStep(requestId, moveHttpRequest, null));
                    requestDictionary.Add(requestId, message.id);
                }

                logger.LogDebug($"Executing batch move request with {requestDictionary.Count} items");
                
                // Execute the batch request
                var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);
                
                // Process responses for each request
                foreach (var kvp in requestDictionary)
                {
                    var requestId = kvp.Key;
                    var messageId = kvp.Value;
                    
                    try
                    {
                        var response = await batchResponse.GetResponseByIdAsync(requestId);
                        var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

                        if ((int)statusCode == 201) // Successful move
                        {
                            successCount++;
                            logger.LogDebug($"Successfully moved message {messageId}");
                            
                            if (response?.Content != null)
                            {
                                var content = await response.Content.ReadAsStringAsync();
                                try
                                {
                                    // Parse the response to get the new message ID
                                    using (JsonDocument document = JsonDocument.Parse(content))
                                    {
                                        if (document.RootElement.TryGetProperty("id", out JsonElement idElement))
                                        {
                                            var newId = idElement.GetString();
                                            var originalMessage = emails.First(m => m.id == messageId);

                                            if (newId != null)
                                            {
                                                originalMessage.newId = newId;
                                                logger.LogDebug($"Message moved. Original ID: {originalMessage.id}, New ID: {newId}");
                                            }
                                        }
                                    }
                                }
                                catch (JsonException ex)
                                {
                                    logger.LogError($"Failed to parse move response for message {messageId}: {ex.Message}", ex);
                                }
                                catch (Exception ex)
                                {
                                    logger.LogError($"Failed to parse move response for message {messageId}: {ex.Message}", ex);
                                }
                            }
                        }
                        else
                        {
                            failureCount++;
                            var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
                            logger.LogWarning($"Failed to move message {messageId}. Status: {(int)statusCode}, Error: {content}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failureCount++;
                        logger.LogError($"Error processing response for message {messageId}: {ex.Message}", ex);
                    }
                }
            }
            catch (ServiceException ex)
            {
                logger.LogError($"Batch move request failed for batch: {ex.Message}", ex);
                failureCount += batch.Count();
                
                // If batch fails, try individual moves as fallback
                logger.LogDebug("Attempting individual moves as fallback");
                foreach (var message in batch)
                {
                    try
                    {
                        var movedMessage = await graphClient.Users[email]
                            .Messages[message.id]
                            .Move
                            .PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                            {
                                DestinationId = "DeletedItems"
                            });
                        
                        if (movedMessage?.Id != null)
                        {
                            message.newId = movedMessage.Id;
                            successCount++;
                            failureCount--;
                            logger.LogDebug($"Individual move successful for message {message.id}");
                        }
                    }
                    catch (Exception individualEx)
                    {
                        logger.LogError($"Individual move failed for message {message.id}: {individualEx.Message}", individualEx);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Unexpected error in batch move: {ex.Message}", ex);
                failureCount += batch.Count();
            }
        }
        
        var batchDuration = DateTime.UtcNow - batchStartTime;
        logger.LogBatchOperation("Move to DeletedItems", emails.Count, successCount, failureCount, batchDuration);
        
        return emails;
    }

    /// <summary>
    /// Permanently deletes emails from DeletedItems folder using hard delete
    /// Processes in batches using Microsoft Graph API batch operations
    /// Falls back to individual deletes if batch operation fails
    /// </summary>
    /// <param name="emails">List of emails to delete (with newId set from move operation)</param>
    /// <param name="graphClient">Authenticated Graph API client</param>
    /// <param name="email">Email address of the mailbox being processed</param>
    private static async Task ProcessDeleteBatchAsync(List<EmailData> emails, GraphServiceClient graphClient, string email)
    {
        logger.LogDebug($"Starting batch delete operation for {emails.Count} emails from DeletedItems");

        var successCount = 0;
        var failureCount = 0;
        var batchStartTime = DateTime.UtcNow;

        // Process emails in smaller batches to avoid timeout issues
        var batches = emails.Chunk(BATCH_SIZE);
        
        foreach (var batch in batches)
        {
            try
            {
                logger.LogDebug($"Processing delete batch of {batch.Count()} emails");
                
                var batchRequestContent = new BatchRequestContentCollection(graphClient);
                var requestDictionary = new Dictionary<string, string>();
                
                foreach (var message in batch)
                {
                    if (string.IsNullOrEmpty(message.newId))
                    {
                        logger.LogWarning($"Skipping delete for message {message.id} - no newId available (move may have failed)");
                        continue;
                    }
                    
                    logger.LogDebug($"Preparing delete request for message ID: {message.newId} (original: {message.id})");
                    
                    var deleteRequest = graphClient.Users[email]
                        .MailFolders["DeletedItems"]
                        .Messages[message.newId]
                        .ToDeleteRequestInformation();

                    // Add headers for hard delete
                    deleteRequest.Headers.Add("Prefer", "permanent");

                    // Convert RequestInformation to HttpRequestMessage
                    var deleteHttpRequest = await graphClient.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
                        deleteRequest,
                        default(CancellationToken)
                    );

                    // Add the request to the batch using the proper method
                    var requestId = Guid.NewGuid().ToString();
                    batchRequestContent.AddBatchRequestStep(new BatchRequestStep(requestId, deleteHttpRequest, null));
                    requestDictionary.Add(requestId, message.newId);
                }

                logger.LogDebug($"Executing batch delete request with {requestDictionary.Count} items");
                
                // Execute the batch request
                var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);
                
                // Process responses for each request
                foreach (var kvp in requestDictionary)
                {
                    var requestId = kvp.Key;
                    var messageId = kvp.Value;
                    
                    try
                    {
                        var response = await batchResponse.GetResponseByIdAsync(requestId);
                        var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

                        if ((int)statusCode == 204) // 204 is success for delete
                        {
                            successCount++;
                            logger.LogDebug($"Successfully deleted message {messageId}");
                        }
                        else
                        {
                            failureCount++;
                            var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
                            logger.LogWarning($"Failed to delete message {messageId}. Status: {(int)statusCode}, Error: {content}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failureCount++;
                        logger.LogError($"Error processing delete response for message {messageId}: {ex.Message}", ex);
                    }
                }
            }
            catch (ServiceException ex)
            {
                logger.LogError($"Batch delete request failed for batch: {ex.Message}", ex);
                failureCount += batch.Count();
                
                // If batch fails, try individual deletes as fallback
                logger.LogDebug("Attempting individual deletes as fallback");
                foreach (var message in batch)
                {
                    if (string.IsNullOrEmpty(message.newId))
                    {
                        logger.LogWarning($"Skipping individual delete for message {message.id} - no newId available");
                        continue;
                    }
                    
                    try
                    {
                        await graphClient.Users[email]
                            .MailFolders["DeletedItems"]
                            .Messages[message.newId]
                            .DeleteAsync();
                        
                        successCount++;
                        failureCount--;
                        logger.LogDebug($"Individual delete successful for message {message.newId}");
                    }
                    catch (Exception individualEx)
                    {
                        logger.LogError($"Individual delete failed for message {message.newId}: {individualEx.Message}", individualEx);
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Unexpected error in batch delete: {ex.Message}", ex);
                failureCount += batch.Count();
            }
        }
        
        var batchDuration = DateTime.UtcNow - batchStartTime;
        logger.LogBatchOperation("Delete from DeletedItems", emails.Count, successCount, failureCount, batchDuration);
    }

    /// <summary>
    /// Exports email metadata to Excel file for audit trail
    /// Creates new file or appends to existing file
    /// Handles file locking by creating timestamped backup files
    /// </summary>
    /// <param name="email">Email address (used for filename)</param>
    /// <returns>True if successful, false otherwise</returns>
    private static bool WriteToExcel(string email)
    {
        try
        {
            logger.LogDebug($"Starting Excel write operation for {email} with {_emails.Count} emails");

            // Sanitize email address for use as filename to prevent path traversal
            string name = SanitizeFileName(email.Split('@')[0]);
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();
            var path = config["ConfigFile:excelDir"];
            
            if (string.IsNullOrEmpty(path))
            {
                logger.LogWarning("Excel directory not configured, using default path");
                path = AppContext.BaseDirectory;
            }
            
            var xlPath = Path.Combine(path, $"{name}.xlsx");
            logger.LogDebug($"Excel file path: {xlPath}");

            XLWorkbook workbook;
            IXLWorksheet worksheet;
            int lastRow = 1; // Default starting row

            // Check if file exists
            if (File.Exists(xlPath))
            {
                logger.LogDebug($"Loading existing Excel file: {xlPath}");
                // Load existing workbook
                workbook = new XLWorkbook(xlPath);
                worksheet = workbook.Worksheet(1);
                lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                logger.LogDebug($"Existing file has {lastRow} rows");
            }
            else
            {
                logger.LogDebug($"Creating new Excel file: {xlPath}");
                // Create new workbook
                workbook = new XLWorkbook();
                worksheet = workbook.Worksheets.Add("Emails");

                // Add headers
                worksheet.Cell(1, 1).Value = "From";
                worksheet.Cell(1, 2).Value = "To";
                worksheet.Cell(1, 3).Value = "Subject";
                worksheet.Cell(1, 4).Value = "ReceivedDateTime";
                worksheet.Cell(1, 5).Value = "Body";
                worksheet.Cell(1, 6).Value = "Source";

                // Set column widths for better readability
                worksheet.Column(1).Width = 30; // From
                worksheet.Column(2).Width = 30; // To
                worksheet.Column(3).Width = 50; // Subject
                worksheet.Column(4).Width = 20; // ReceivedDateTime
                worksheet.Column(5).Width = 100; // Body
                worksheet.Column(6).Width = 20; // Source

                // Optional: Add some basic formatting to headers
                var headerRow = worksheet.Row(1);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;
                logger.LogDebug("Excel headers created successfully");
            }

            // Add new data starting from the row after the last used row
            logger.LogDebug($"Adding {_emails.Count} email records to Excel starting from row {lastRow + 1}");
            
            for (int i = 0; i < _emails.Count; i++)
            {
                var row = lastRow + i + 1;
                worksheet.Cell(row, 1).Value = _emails[i].from;
                worksheet.Cell(row, 2).Value = string.Join(", ", _emails[i].toRecipients);
                worksheet.Cell(row, 3).Value = _emails[i].subject;
                worksheet.Cell(row, 4).Value = _emails[i].receivedDateTime;
                worksheet.Cell(row, 5).Value = _emails[i].body;
                worksheet.Cell(row, 6).Value = _emails[i].source;
            }
            
            logger.LogDebug($"Added {_emails.Count} email records to Excel");

            // Auto-fit columns based on content
            worksheet.Columns().AdjustToContents();
            logger.LogDebug("Excel columns auto-fitted");

            // Ensure the directory exists
            Directory.CreateDirectory(path);
            logger.LogDebug($"Ensured directory exists: {path}");

            // Try to save the file
            try
            {
                logger.LogDebug($"Attempting to save Excel file: {xlPath}");
                
                // If file exists, we need to ensure it's not locked
                if (File.Exists(xlPath))
                {
                    workbook.Save();
                    logger.LogDebug("Excel file saved successfully (existing file)");
                }
                else
                {
                    workbook.SaveAs(xlPath);
                    logger.LogDebug("Excel file saved successfully (new file)");
                }
            }
            catch (IOException ex)
            {
                // If file is locked, try saving with a timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                string alternativePath = Path.Combine(path, $"{name}_{timestamp}.xlsx");
                workbook.SaveAs(alternativePath);
                logger.LogWarning($"Original file was locked. Saved to alternative path: {alternativePath}");
                logger.LogError($"File lock error details", ex);
            }
            finally
            {
                workbook.Dispose(); // Ensure workbook is properly disposed
            }

            return true;
        }
        catch (Exception ex)
        {
            logger.LogError($"Error writing to excel file for {email}: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Reads email account configuration from Excel file
    /// Excel file path is specified in appsettings.json
    /// </summary>
    /// <returns>List of ConfigData objects with email account processing rules</returns>
    static List<ConfigData> ReadConfig()
    {
        try
        {
            logger.LogDebug("Starting configuration read");
            
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory) // Set the base path
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Add the JSON config file
                    .Build();
            var excel = config["ConfigFile:path"];
            
            if (string.IsNullOrEmpty(excel))
            {
                logger.LogError("Config file path not found in appsettings.json");
                throw new Exception("Config file not found");
            }
            
            logger.LogDebug($"Config file path: {excel}");
            
            if (!File.Exists(excel))
            {
                logger.LogError($"Config file does not exist: {excel}");
                throw new Exception($"Config file does not exist: {excel}");
            }
        using (var workbook = new XLWorkbook(excel))
        {
            logger.LogDebug("Excel workbook loaded successfully");
            
            var arr = new List<ConfigData>();
            // Get the first worksheet
            var worksheet = workbook.Worksheet(1);
            logger.LogDebug("First worksheet accessed");

            // Read data from the worksheet
            var lastRow = worksheet.LastRowUsed()?.RowNumber();
            logger.LogDebug($"Found {lastRow} rows in configuration file");
            
            int iInbox = 0, iDeleted = 0, iSent = 0, iAttachment = 0, iBody = 0, iRead = 0;
            int processedRows = 0;

            for (int row = 2; row <= lastRow; row++)
            {
                iInbox = DEFAULT_DAYS_THRESHOLD;
                iDeleted = DEFAULT_DAYS_THRESHOLD;
                iSent = DEFAULT_DAYS_THRESHOLD;
                iAttachment = DEFAULT_ATTACHMENT_FLAG;
                iBody = DEFAULT_BODY_FLAG;
                var email = worksheet.Cell(row, 1).Value;
                var inbox = int.TryParse(worksheet.Cell(row, 2).Value.ToString(), out iInbox);
                var deleted = int.TryParse(worksheet.Cell(row, 3).Value.ToString(), out iDeleted);
                var sent = int.TryParse(worksheet.Cell(row, 4).Value.ToString(), out iSent);
                var body = int.TryParse(worksheet.Cell(row, 5).Value.ToString(), out iBody);
                var attachment = int.TryParse(worksheet.Cell(row, 6).Value.ToString(), out iAttachment);
                var read = int.TryParse(worksheet.Cell(row, 7).Value.ToString(), out iRead);

                var emailAddress = email.ToString();

                // Validate email address
                if (!IsValidEmail(emailAddress))
                {
                    logger.LogWarning($"Skipping row {row}: Invalid email address '{emailAddress}'");
                    continue;
                }

                // Validate days thresholds (must be positive)
                if (iInbox < 0 || iDeleted < 0 || iSent < 0)
                {
                    logger.LogWarning($"Skipping row {row}: Negative days threshold not allowed (Inbox={iInbox}, Deleted={iDeleted}, Sent={iSent})");
                    continue;
                }

                var configData = new ConfigData
                {
                    email = emailAddress,
                    inbox = iInbox,
                    deleted = iDeleted,
                    sent = iSent,
                    body = iBody == 1 ? true : false,
                    attachment = iAttachment == 1 ? true : false,
                    read = iRead == 1 ? true : false
                };

                arr.Add(configData);
                processedRows++;
                
                logger.LogDebug($"Processed row {row}: Email={configData.email}, Inbox={configData.inbox}, Sent={configData.sent}, Deleted={configData.deleted}, Body={configData.body}, Attachment={configData.attachment}, Read={configData.read}");
            }
            
            logger.LogInfo($"Configuration loaded successfully. Processed {processedRows} email accounts");
            return arr;
        }
        }
        catch (Exception ex)
        {
            logger.LogError("Failed to read configuration", ex);
            throw;
        }
    }

    /// <summary>
    /// Extracts plain text from HTML email body
    /// Decodes HTML entities and strips all HTML tags
    /// </summary>
    /// <param name="html">HTML content from email body</param>
    /// <returns>Plain text without HTML formatting</returns>
    private static string ExtractPlainTextFromHtml(string html)
    {
        if (string.IsNullOrEmpty(html))
            return string.Empty;

        // Use HtmlAgilityPack to parse and extract text
        var htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(html);

        // Extract plain text by removing all HTML tags
        // HtmlAgilityPack automatically decodes HTML entities
        var plainText = htmlDoc.DocumentNode.InnerText;

        // Decode any remaining HTML entities using the built-in decoder
        return System.Net.WebUtility.HtmlDecode(plainText);
    }

    /// <summary>
    /// Sanitizes a filename by removing invalid characters and path traversal sequences
    /// Prevents path traversal attacks
    /// </summary>
    /// <param name="fileName">The filename to sanitize</param>
    /// <returns>Sanitized filename safe for use in file system operations</returns>
    private static string SanitizeFileName(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            return "unknown";

        // Remove any path traversal sequences
        fileName = fileName.Replace("..", "").Replace("/", "").Replace("\\", "");

        // Remove invalid filename characters
        var invalidChars = Path.GetInvalidFileNameChars();
        foreach (var c in invalidChars)
        {
            fileName = fileName.Replace(c.ToString(), "");
        }

        // Ensure filename is not empty after sanitization
        if (string.IsNullOrWhiteSpace(fileName))
            return "unknown";

        // Limit filename length to prevent issues with path length limits
        if (fileName.Length > 100)
            fileName = fileName.Substring(0, 100);

        return fileName;
    }

    /// <summary>
    /// Validates an email address format
    /// </summary>
    /// <param name="email">Email address to validate</param>
    /// <returns>True if valid, false otherwise</returns>
    private static bool IsValidEmail(string? email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return false;

        try
        {
            var addr = new System.Net.Mail.MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }
}

/// <summary>
/// Configuration data for an email account
/// Specifies which emails to process and how
/// </summary>
class ConfigData
{
    /// <summary>Email address to process</summary>
    public string? email { get; set; }

    /// <summary>Age threshold in days for Inbox folder</summary>
    public int inbox { get; set; }

    /// <summary>Age threshold in days for DeletedItems folder</summary>
    public int deleted { get; set; }

    /// <summary>Age threshold in days for SentItems folder</summary>
    public int sent { get; set; }

    /// <summary>Include email body in Excel export</summary>
    public bool body { get; set; }

    /// <summary>Only process emails with attachments</summary>
    public bool attachment { get; set; }

    /// <summary>Only process read emails</summary>
    public bool read { get; set; }
}

/// <summary>
/// Email metadata for processing and export
/// Stores information about emails to be deleted
/// </summary>
class EmailData
{
    /// <summary>Original message ID in source folder</summary>
    public string? id { get; set; }

    /// <summary>Sender email address</summary>
    public string? from { get; set; }

    /// <summary>List of recipient email addresses</summary>
    public List<string>? toRecipients { get; set; }

    /// <summary>Email body content (plain text)</summary>
    public string? body { get; set; }

    /// <summary>Email subject line</summary>
    public string? subject { get; set; }

    /// <summary>Date and time email was received (ISO 8601 format)</summary>
    public string? receivedDateTime { get; set; }

    /// <summary>New message ID after moving to DeletedItems</summary>
    public string? newId { get; set; }

    /// <summary>Source folder name (Inbox, SentItems, DeletedItems)</summary>
    public string? source { get; set; }
}




