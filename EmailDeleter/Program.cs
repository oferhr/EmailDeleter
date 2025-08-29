using Azure.Identity;
using ClosedXML.Excel;
using EmailDeleter;
using HtmlAgilityPack;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System.Text.Json;

class Program
{
    private static List<EmailData> _emails = new List<EmailData>();
    private static SimpleLogger logger;
    private static SimpleLogger infoLogger;
    private static Dictionary<string, int> counts = new Dictionary<string, int>();
    

    static async Task Main(string[] args)
    {
        var startTime = DateTime.UtcNow;
        
        // Initialize logger first
        InitializeLogger();
        logger.LogInfo("EmailDeleter application started");
        
        try
        {
            setInfoLogger();
            logger.LogInfo("Info logger initialized successfully");
            
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
            }
            
            var totalDuration = DateTime.UtcNow - startTime;
            logger.LogPerformance("Total application execution", totalDuration, $"Total emails processed: {counts.Values.Sum()}");
            infoLogger.LogInfo($"Emails deleted: Inbox: {counts["Inbox"]}, SentItems: {counts["SentItems"]}, DeletedItems: {counts["DeletedItems"]}");
            logger.LogInfo("EmailDeleter application completed successfully");
        }
        catch (Exception ex)
        {
            logger.LogError("Application failed with unexpected error", ex);
            throw;
        }
    }
    
    private static void InitializeLogger()
    {
        try
        {
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();
            

            var enableDebugLogging = config.GetSection("Logging").Value != null
                ? bool.Parse(config.GetSection("Logging").Value)
                : false;
            logger = new SimpleLogger("log", string.Empty, enableDebugLogging);
            
            Console.WriteLine($"Logger initialized with debug logging: {enableDebugLogging}");
        }
        catch (Exception ex)
        {
            // Fallback to basic logger if configuration fails
            logger = new SimpleLogger("log", string.Empty, false);
            Console.WriteLine($"Failed to initialize logger with configuration, using fallback: {ex.Message}");
        }
    }
    
    private static void setInfoLogger()
    {
        try
        {
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory) // Set the base path
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Add the JSON config file
                    .Build();
            var path = config["CofigFile:infoLogDir"];
            if (string.IsNullOrEmpty(path))
            {
                logger.LogWarning("Info log directory not configured, using default path");
                path = AppContext.BaseDirectory;
            }
            infoLogger = new SimpleLogger("log", path, true);
            logger.LogInfo($"Info logger initialized with path: {path}");
        }
        catch (Exception ex)
        {
            logger.LogError("Failed to initialize info logger", ex);
            throw;
        }
    }
    static async Task<int> fetchEmails(ConfigData config, string dir, int days, bool isDeleted)
    {
        var startTime = DateTime.UtcNow;
        logger.LogInfo($"Starting email fetch for {config.email} in {dir} folder (days: {days}, isDeleted: {isDeleted})");

        // Load secrets from graph-secrets.json
        // We keep secrets in external file to avoid committing them to git during development
        // for production use 
        // var clientId = "xxx";
        // var tenentId = "xxx";
        // var secret = "xxx";
        var secretsConfig = new ConfigurationBuilder()
           .SetBasePath(AppContext.BaseDirectory)
           .AddJsonFile("graph-secrets.json", optional: false, reloadOnChange: true)
           .Build();
        var clientId = secretsConfig["clientId"];
        var tenentId = secretsConfig["tenantId"];
        var secret = secretsConfig["secret"];

        //var clientId = "xxx";
        //var tenentId = "xxx";
        //var secret = "xxx";.

        logger.LogDebug($"Using tenant ID: {tenentId}, client ID: {clientId}");
        infoLogger.LogInfo($"Fetching emails for {config.email} in {dir} folder.");
        
        var clientSecretCredential = new ClientSecretCredential(tenentId, clientId, secret);
        var graphClient = new GraphServiceClient(clientSecretCredential);
        int counter = 0;
        try
        {
            var dateThreshold = DateTime.Now.AddDays(days*-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
            var attachment = config.attachment ? "hasAttachments eq true and" : "";
            var filter = $"{attachment} receivedDateTime gt {dateThreshold}";
            
            logger.LogDebug($"Using filter: {filter}");
            logger.LogDebug($"Date threshold: {dateThreshold}");
            
            var page = await graphClient.Users[config.email]
                .MailFolders[dir]
                .Messages
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = filter;
                    requestConfiguration.QueryParameters.Select = new string[] { "subject,body,receivedDateTime,from,toRecipients,isRead" };
                    requestConfiguration.QueryParameters.Top = 10;
                });


            
            var pageCount = 0;
            while (page != null)
            {
                pageCount++;
                logger.LogDebug($"Processing page {pageCount} with {page.Value?.Count ?? 0} messages");
                
                foreach (var message in page.Value)
                {
                    logger.LogDebug($"Processing message ID: {message.Id}, Subject: {message.Subject?.Substring(0, Math.Min(50, message.Subject.Length))}...");
                    
                    var htmlBody = message.Body?.Content ?? "No content available.";
                    var plainTextBody = ExtractPlainTextFromHtml(htmlBody);
                    var emailData = new EmailData
                    {
                        from = message.From?.EmailAddress?.Address ?? "Unknown",
                        toRecipients = message.ToRecipients?.Select(r => r.EmailAddress?.Address).ToList() ?? new List<string>(),
                        subject = message.Subject ?? "No Subject",
                        receivedDateTime = message.ReceivedDateTime?.ToString("yyyy-MM-ddTHH:mm:ssZ") ?? "",
                        body = plainTextBody.Trim(),
                        id = message.Id
                    };

                    _emails.Add(emailData);
                    

                    //Console.WriteLine($"Processing email...");
                    //Console.WriteLine($"Subject: {emailData.subject}");
                    //Console.WriteLine($"From: {emailData.from}");
                    //Console.WriteLine($"To: {string.Join(", ", emailData.toRecipients)}");
                    //Console.WriteLine($"Received on: {emailData.receivedDateTime}");
                    //Console.WriteLine($"Body: {emailData.body}");
                    //Console.WriteLine(new string('-', 50));

                    //Delete the email
                    //await graphClient.Users[config.email]
                    //    .Messages[message.Id]
                    //    .DeleteAsync(Microsoft.Graph.Models.DeletionMode.HardDelete);

                    // Move the message to the Deleted Items folder


                    //Console.WriteLine("Email deleted.");
                }
                var isOK = true;
                var newEmails = new List<EmailData>();
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
                    logger.LogDebug($"Skipping move operation for DeletedItems folder");
                    newEmails = _emails;
                }
                
                if (isOK)
                {
                    logger.LogDebug($"Processing delete batch for {newEmails.Count} messages");
                    var deleteStartTime = DateTime.UtcNow;
                    await ProcessDeleteBatchAsync(newEmails, graphClient, config.email);
                    var deleteDuration = DateTime.UtcNow - deleteStartTime;
                    logger.LogPerformance($"Delete batch for {newEmails.Count} messages", deleteDuration);
                }
                //var isok = await moveToDeleted(_emails, graphClient, config.email);
                //await ProcessDeleteBatchAsync(_emails, graphClient, config.email);
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

            //Console.WriteLine("All emails processed and deleted.");
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
    static async Task<List<EmailData>> moveToDeleted(List<EmailData> emails, GraphServiceClient graphClient, string email)
    {
        logger.LogDebug($"Starting batch move operation for {emails.Count} emails to DeletedItems");
        
        var batchRequestContent = new BatchRequestContentCollection(graphClient);
        var requestDictionary = new Dictionary<string, RequestInformation>();
        var messageIdMapping = new Dictionary<string, string>();
        
        foreach (var message in emails)
        {
            logger.LogDebug($"Preparing move request for message ID: {message.id}");
            
            var moveRequest = graphClient.Users[email]
                .Messages[message.id]
                .Move
                .ToPostRequestInformation(new Microsoft.Graph.Users.Item.Messages.Item.Move.MovePostRequestBody
                {
                    DestinationId = "DeletedItems"
                });

            var httpRequest = await graphClient.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
                moveRequest,
                default(CancellationToken)
            );

            // Add request to batch
            var requestId = message.id;
            var batchRequestStep = new BatchRequestStep(requestId, httpRequest, null);
            batchRequestContent.AddBatchRequestStep(batchRequestStep);
            requestDictionary.Add(requestId, moveRequest);
        }

        // Execute the batch request
        try
        {
            logger.LogDebug($"Executing batch move request with {requestDictionary.Count} items");
            var batchStartTime = DateTime.UtcNow;
            
            // Execute the batch request
            var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);
            
            var batchDuration = DateTime.UtcNow - batchStartTime;
            logger.LogPerformance($"Batch move request execution", batchDuration);

            var successCount = 0;
            var failureCount = 0;
            
            foreach (var requestId in requestDictionary.Keys)
            {
                var response = await batchResponse.GetResponseByIdAsync(requestId);
                var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

                if ((int)statusCode == 201) // Successful move
                {
                    successCount++;
                    logger.LogDebug($"Successfully moved message {requestId}");
                    
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
                                    var originalMessage = emails.First(m => m.id == requestId);

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
                            logger.LogError($"Failed to parse move response for message {requestId}: {ex.Message}", ex);
                        }
                        catch (Exception ex)
                        {
                            logger.LogError($"Failed to parse move response for message {requestId}: {ex.Message}", ex);
                        }
                    }
                }
                else
                {
                    failureCount++;
                    var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
                    logger.LogWarning($"Failed to move message {requestId}. Status: {(int)statusCode}, Error: {content}");
                }
            }
            
            logger.LogBatchOperation("Move to DeletedItems", requestDictionary.Count, successCount, failureCount, batchDuration);
            // Process responses for each request
            //foreach (var requestId in requestDictionary.Keys)
            //{
            //    var response = await batchResponse.GetResponseByIdAsync(requestId);
            //    var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

            //    if ((int)statusCode != 201) // 201 is success for move operation
            //    {
            //        var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
            //        infoLogger.LogInfo($"Failed to move message {requestId}. Status: {(int)statusCode}, Error: {content}");
            //        //  Console.WriteLine($"Failed to move message {requestId}. Status: {(int)statusCode}, Error: {content}");
            //    }
            //}
        }
        catch (ServiceException ex)
        {
            logger.LogError($"Batch move request failed: {ex.Message}", ex);
            //Console.WriteLine($"Batch move request failed: {ex.Message}");
            throw;
        }
        return emails;
    }
    private static async Task ProcessDeleteBatchAsync(List<EmailData> emails, GraphServiceClient graphClient, string email)
    {
        logger.LogDebug($"Starting batch delete operation for {emails.Count} emails from DeletedItems");
        
        var batchRequestContent = new BatchRequestContentCollection(graphClient);
        var requestDictionary = new Dictionary<string, RequestInformation>();

        foreach (var message in emails)
        {
            logger.LogDebug($"Preparing delete request for message ID: {message.newId} (original: {message.id})");
            
            var deleteRequest = graphClient.Users[email]
                .MailFolders["DeletedItems"]
                .Messages[message.newId]
                .ToDeleteRequestInformation();

            // Add headers for hard delete
            deleteRequest.Headers.Add("Prefer", "permanent");

            // Convert RequestInformation to HttpRequestMessage
            var httpRequest = await graphClient.RequestAdapter.ConvertToNativeRequestAsync<HttpRequestMessage>(
                deleteRequest,
                default(CancellationToken)
            );

            logger.LogDebug($"Delete request URL: {httpRequest.RequestUri}");
            logger.LogDebug($"Delete request method: {httpRequest.Method}");

            // Add request to batch
            var requestId = Guid.NewGuid().ToString();
            var batchRequestStep = new BatchRequestStep(requestId, httpRequest, null);
            batchRequestContent.AddBatchRequestStep(batchRequestStep);
            requestDictionary.Add(requestId, deleteRequest);
        }

        try
        {
            logger.LogDebug($"Executing batch delete request with {requestDictionary.Count} items");
            var batchStartTime = DateTime.UtcNow;
            
            // Execute the batch request
            var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);
            
            var batchDuration = DateTime.UtcNow - batchStartTime;
            logger.LogPerformance($"Batch delete request execution", batchDuration);

            var successCount = 0;
            var failureCount = 0;
            
            // Process responses for each request
            foreach (var requestId in requestDictionary.Keys)
            {
                var response = await batchResponse.GetResponseByIdAsync(requestId);
                var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

                if ((int)statusCode == 204) // 204 is success for delete
                {
                    successCount++;
                    logger.LogDebug($"Successfully deleted message {requestId}");
                }
                else
                {
                    failureCount++;
                    var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
                    logger.LogWarning($"Failed to delete message {requestId}. Status: {(int)statusCode}, Error: {content}");
                }
            }
            
            logger.LogBatchOperation("Delete from DeletedItems", requestDictionary.Count, successCount, failureCount, batchDuration);
        }
        catch (ServiceException ex)
        {
            logger.LogError($"Batch delete request failed: {ex.Message}", ex);
           // Console.WriteLine($"Batch delete request failed: {ex.Message}");
            throw;
        }
        catch (Exception ex)
        {
            logger.LogError($"Batch delete request failed: {ex.Message}", ex);
           // Console.WriteLine($"Batch delete request failed: {ex.Message}");
            throw;
        }
    }
    private static bool WriteToExcel(string email)
    {
        try
        {
            logger.LogDebug($"Starting Excel write operation for {email} with {_emails.Count} emails");
            
            string name = email.Split('@')[0];
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();
            var path = config["CofigFile:excelDir"];
            
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

                // Set column widths for better readability
                worksheet.Column(1).Width = 30; // From
                worksheet.Column(2).Width = 30; // To
                worksheet.Column(3).Width = 50; // Subject
                worksheet.Column(4).Width = 20; // ReceivedDateTime
                worksheet.Column(5).Width = 100; // Body

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
        finally
        {
            // Ensure proper cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
    static List<ConfigData> ReadConfig()
    {
        try
        {
            logger.LogDebug("Starting configuration read");
            
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory) // Set the base path
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Add the JSON config file
                    .Build();
            var excel = config["CofigFile:path"];
            
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
                iInbox = 30; iDeleted = 30; iSent = 30; iAttachment = 1; iBody = 1;
                var email = worksheet.Cell(row, 1).Value;
                var inbox = int.TryParse(worksheet.Cell(row, 2).Value.ToString(), out iInbox);
                var deleted = int.TryParse(worksheet.Cell(row, 3).Value.ToString(), out iDeleted);
                var sent = int.TryParse(worksheet.Cell(row, 4).Value.ToString(), out iSent);
                var body = int.TryParse(worksheet.Cell(row, 5).Value.ToString(), out iBody);
                var attachment = int.TryParse(worksheet.Cell(row, 6).Value.ToString(), out iAttachment);
                var read = int.TryParse(worksheet.Cell(row, 7).Value.ToString(), out iRead);

                var configData = new ConfigData
                {
                    email = email.ToString(),
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
    private static string ExtractPlainTextFromHtml(string html)
    {
        if (string.IsNullOrEmpty(html))
            return string.Empty;

        var htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(html);

        // Extract plain text by removing all HTML tags
        return htmlDoc.DocumentNode.InnerText;
    }
}

class ConfigData
{
    public string? email { get; set; }
    public int inbox { get; set; }
    public int deleted { get; set; }
    public int sent { get; set; }
    public bool body { get; set; }
    public bool attachment { get; set; }
    public bool read { get; set; }
}
class EmailData
{
    public string? id { get; set; }
    public string? from { get; set; }
    public List<string>? toRecipients { get; set; }
    public string? body { get; set; }
    public string? subject { get; set; }
    public string? receivedDateTime { get; set; }
    public string? newId { get; set; }
}




