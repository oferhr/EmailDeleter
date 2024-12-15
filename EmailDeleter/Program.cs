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
    static List<EmailData> _emails = new List<EmailData>();
    static SimpleLogger logger = new SimpleLogger("log", string.Empty);
    static SimpleLogger infoLogger;
    static Dictionary<string, int> counts = new Dictionary<string, int>();
    

    static async Task Main(string[] args)
    {

        setInfoLogger();
        var ConfigData = ReadConfig();
        foreach (var config in ConfigData)
        {
            counts["Inbox"] =  await fetchEmails(config, "Inbox", config.inbox, false);
            counts["SentItems"] = await fetchEmails(config, "SentItems", config.sent, false);
            counts["DeletedItems"] = await fetchEmails(config, "DeletedItems", config.deleted, true);
        }
        infoLogger.LogInfo($"Emails deleted: Inbox: {counts["Inbox"]}, SentItems: {counts["SentItems"]}, DeletedItems: {counts["DeletedItems"]}");
    }
    private static void setInfoLogger()
    {
        var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory) // Set the base path
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Add the JSON config file
                .Build();
        var path = config["CofigFile:infoLogDir"];
        infoLogger = new SimpleLogger("log", path);
    }
    static async Task<int> fetchEmails(ConfigData config, string dir, int days, bool isDeleted)
    {
        //var clientId = "14ef1b2a-da52-4aec-8438-0a5c24326c6e";
        //var tenentId = "7a5b4a33-c120-4e8e-a41b-e5907534e115";
        //var secret = "R0B8Q~ShoSVlB_w2tY~ruRXHBfYOUkF.3Z-WIaxx";


        var clientId = "dd85b611-33c5-4744-8ef8-56d389a792bb";
        var tenentId = "ca91d845-398d-4116-a8a7-c23eabe5d9a7";
        var secret = "C1e8Q~9vnzNXL5wVnPzvusC9f3DNOV.L5IGzYcVf";

        infoLogger.LogInfo($"Fetching emails for {config.email} in {dir} folder.");
        var clientSecretCredential = new ClientSecretCredential(tenentId, clientId, secret);
        var graphClient = new GraphServiceClient(clientSecretCredential);
        int counter = 0;
        try
        {
            var dateThreshold = DateTime.UtcNow.AddDays(days*-1).ToString("yyyy-MM-ddTHH:mm:ssZ");
            var attachment = config.attachment ? "hasAttachments eq true and" : "";
            var page = await graphClient.Users[config.email]
                .MailFolders[dir]
                .Messages
                .GetAsync(requestConfiguration =>
                {
                    requestConfiguration.QueryParameters.Filter = $"{attachment} receivedDateTime gt {dateThreshold}";
                    requestConfiguration.QueryParameters.Select = new string[] { "subject,body,receivedDateTime,from,toRecipients,isRead" };
                    requestConfiguration.QueryParameters.Top = 10;
                });


            
            while (page != null)
            {
                foreach (var message in page.Value)
                {
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
                        newEmails = await moveToDeleted(_emails, graphClient, config.email);
                        //foreach (var mapping in newIds)
                        //{
                        //    messageIdMapping[mapping.Key] = mapping.Value;
                        //}
                        //var messagesWithNewIds = _emails.Select(msg => new Message
                        //{
                        //    Id = messageIdMapping.GetValueOrDefault(msg.id, msg.id)
                        //}).ToList();
                    }
                    catch (Exception ex)
                    {
                        logger.LogError($"Error fetching messages: {ex.Message}", ex);
                        isOK = false;
                    }
                    
                }
                
                if (isOK)
                {
                    await ProcessDeleteBatchAsync(newEmails, graphClient, config.email);
                }
                //var isok = await moveToDeleted(_emails, graphClient, config.email);
                //await ProcessDeleteBatchAsync(_emails, graphClient, config.email);
                WriteToExcel(config.email);
                infoLogger.LogInfo($"{_emails.Count.ToString()} Emails deleted for {config.email} in {dir} folder.");
                counter += _emails.Count;
                _emails = new List<EmailData>();
                // Fetch the next page using the OdataNextLink
                if (!string.IsNullOrEmpty(page.OdataNextLink))
                {
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
                    page = null; // No more pages
                }
            }

            //Console.WriteLine("All emails processed and deleted.");
        }
        catch (ServiceException ex)
        {
            logger.LogError($"Error fetching messages: {ex.Message}", ex);
            //Console.WriteLine($"Error fetching messages: {ex.Message}");
        }
        catch (Exception ex)
        {
            logger.LogError($"Error fetching messages: {ex.Message}", ex);
            //Console.WriteLine($"Error: {ex.Message}");*/
        }
        
        return counter;
    }
    static async Task<List<EmailData>> moveToDeleted(List<EmailData> emails, GraphServiceClient graphClient, string email)
    {
        var batchRequestContent = new BatchRequestContentCollection(graphClient);
        var requestDictionary = new Dictionary<string, RequestInformation>();
        var messageIdMapping = new Dictionary<string, string>();
        //List<EmailData> _emails = new List<EmailData>();
        foreach (var message in emails)
        {
            
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
            // Execute the batch request
            var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);

            foreach (var requestId in requestDictionary.Keys)
            {
                var response = await batchResponse.GetResponseByIdAsync(requestId);
                var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

                if ((int)statusCode == 201) // Successful move
                {
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
                                        //Console.WriteLine($"Message moved. Original ID: {originalMessage.id}, New ID: {newId}");
                                    }
                                }
                            }
                        }
                        catch (JsonException ex)
                        {
                            //Console.WriteLine($"Failed to parse move response: {ex.Message}");
                            logger.LogError($"Failed to parse move response: {ex.Message}", ex);
                        }
                        catch (Exception ex)
                        {
                            //Console.WriteLine($"Failed to parse move response: {ex.Message}");
                            logger.LogError($"Failed to parse move response: {ex.Message}", ex);
                        }
                    }
                }
                else
                {
                    var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
                    infoLogger.LogInfo($"Failed to move message. Status: {(int)statusCode}, Error: {content}");
                }
            }
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
        var batchRequestContent = new BatchRequestContentCollection(graphClient);
        var requestDictionary = new Dictionary<string, RequestInformation>();

        foreach (var message in emails)
        {
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

            //Console.WriteLine($"Delete request URL: {httpRequest.RequestUri}");
            //Console.WriteLine($"Delete request method: {httpRequest.Method}");
            //Console.WriteLine($"Delete request headers: {string.Join(", ", httpRequest.Headers)}");

            //// Try to get the message first to verify it exists
            //try
            //{
            //    var messageCheck = await graphClient.Users[email]
            //        .MailFolders["DeletedItems"]
            //        .Messages[message.newId]
            //        .GetAsync();
            //    Console.WriteLine($"Message found in DeletedItems: {message.newId}");
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"Failed to find message {message.newId} in DeletedItems: {ex.Message}");
            //}

            // Add request to batch
            var requestId = Guid.NewGuid().ToString();
            var batchRequestStep = new BatchRequestStep(requestId, httpRequest, null);
            batchRequestContent.AddBatchRequestStep(batchRequestStep);
            requestDictionary.Add(requestId, deleteRequest);
        }

        try
        {
            // Execute the batch request
            var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);

            // Process responses for each request
            foreach (var requestId in requestDictionary.Keys)
            {
                var response = await batchResponse.GetResponseByIdAsync(requestId);
                var statusCode = response?.StatusCode ?? System.Net.HttpStatusCode.NotFound;

                if ((int)statusCode != 204) // 204 is success for delete
                {
                    var content = response != null ? await response.Content.ReadAsStringAsync() : "No content";
                    infoLogger.LogInfo($"Failed to delete message {requestId}. Status: {(int)statusCode}, Error: {content}");
                }
            }
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
            string name = email.Split('@')[0];
            var config = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                    .Build();
            var path = config["CofigFile:excelDir"];
            var xlPath = Path.Combine(path, $"{name}.xlsx");

            XLWorkbook workbook;
            IXLWorksheet worksheet;
            int lastRow = 1; // Default starting row

            // Check if file exists
            if (File.Exists(xlPath))
            {
                // Load existing workbook
                workbook = new XLWorkbook(xlPath);
                worksheet = workbook.Worksheet(1);
                lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
            }
            else
            {
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
            }

            // Add new data starting from the row after the last used row
            for (int i = 0; i < _emails.Count; i++)
            {
                var row = lastRow + i + 1;
                worksheet.Cell(row, 1).Value = _emails[i].from;
                worksheet.Cell(row, 2).Value = string.Join(", ", _emails[i].toRecipients);
                worksheet.Cell(row, 3).Value = _emails[i].subject;
                worksheet.Cell(row, 4).Value = _emails[i].receivedDateTime;
                worksheet.Cell(row, 5).Value = _emails[i].body;

            }

            // Auto-fit columns based on content
            worksheet.Columns().AdjustToContents();

            // Ensure the directory exists
            Directory.CreateDirectory(path);

            // Try to save the file
            try
            {
                // If file exists, we need to ensure it's not locked
                if (File.Exists(xlPath))
                {
                    workbook.Save();
                }
                else
                {
                    workbook.SaveAs(xlPath);
                }
            }
            catch (IOException ex)
            {
                // If file is locked, try saving with a timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                string alternativePath = Path.Combine(path, $"{name}_{timestamp}.xlsx");
                workbook.SaveAs(alternativePath);
                logger.LogError($"Original file was locked. Saved to alternative path: {alternativePath}", ex);
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
        var config = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory) // Set the base path
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true) // Add the JSON config file
                .Build();
        var excel = config["CofigFile:path"];
        if (string.IsNullOrEmpty(excel))
        {
            throw new Exception("Config file not found");
        }
        using (var workbook = new XLWorkbook(excel))
        {
            var arr = new List<ConfigData>();
            // Get the first worksheet
            var worksheet = workbook.Worksheet(1);

            // Read data from the worksheet
            var lastRow = worksheet.LastRowUsed()?.RowNumber();
            int iInbox = 0, iDeleted = 0, iSent = 0, iAttachment = 0, iBody = 0, iRead = 0;
            
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

                arr.Add(new ConfigData
                {
                    email = email.ToString(),
                    inbox = iInbox,
                    deleted = iDeleted,
                    sent = iSent,
                    body = iBody == 1 ? true : false,
                    attachment = iAttachment == 1 ? true : false,
                    read = iRead == 1 ? true : false
                });
            }
            return arr;
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




