using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System.IO;

namespace MSGraphAPI_LargeAttachment
{
    class Program
    {
        private static GraphServiceClient _graphServiceClient;
        private static HttpClient _httpClient;

        private static string userId;

        static void Main(string[] args)
        {
            // Load appsettings.json
            var config = LoadAppSettings();
            if (null == config)
            {
                Console.WriteLine("Missing or invalid appsettings.json file.");
                return;
            }

            // Query using Graph SDK (preferred when possible)
            GraphServiceClient graphClient = GetAuthenticatedGraphClient(config);
            HttpClient httpClient = GetAuthenticatedHTTPClient(config);

            // Create message
            var message = CreateMessage();

            // Attachment
            var attachmentName = "Sample.pdf";
            var attachmentStream = new FileStream(attachmentName, FileMode.Open, FileAccess.Read);
            var attachmentSize = attachmentStream.Length;

            // Create upload session
            var uploadSession = CreateUploadSession(message, attachmentName, attachmentSize);

            // Upload attachment
            var uploadResult = UploadAttachment(uploadSession, attachmentStream);

            // Send message
            SendMessage(message);
        }

        // Create message
        private static Message CreateMessage()
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);

            var message = new Message
            {
                Subject = "Large attachment is here",
                Importance = Importance.Low,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "Why can't you make small <b>attachments</b>!"
                },
                ToRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Address = "someone@replacehere.com"
                        }
                    }
                }
            };

            message = mailboxHelper.CreateMessage(message, userId).Result;
            Console.WriteLine("Message Id: " + message.Id);

            return message;
        }

        // Create upload session
        private static UploadSession CreateUploadSession(Message message, string attachmentName, long attachmentSize)
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);

            var attachmentItem = new AttachmentItem()
            {
                AttachmentType = AttachmentType.File,
                Name = attachmentName,
                Size = attachmentSize
            };

            var uploadSession = mailboxHelper.CreateUploadSession(message, userId, attachmentItem).Result;
            Console.WriteLine("Upload URL: " + uploadSession.UploadUrl);

            return uploadSession;
        }

        // Upload attachment
        private static UploadResult<AttachmentItem> UploadAttachment(UploadSession uploadSession, FileStream attachmentStream)
        {
            var mailboxHelper = new MailboxHelper(_httpClient);

            var uploadResult = mailboxHelper.UploadAttachment(uploadSession, attachmentStream).Result;
            Console.WriteLine("Upload Succeeded: " + uploadResult.UploadSucceeded);

            return uploadResult;
        }

        // Send Message
        private static Message SendMessage(Message message)
        {
            var mailboxHelper = new MailboxHelper(_graphServiceClient);

            mailboxHelper.SendMessage(message, userId);
            Console.WriteLine("Message sent");

            return message;
        }



        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _httpClient = new HttpClient(new AuthHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];

            userId = config["userId"];

            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["tenantId"]) ||
                    string.IsNullOrEmpty(config["userId"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }
    }
}
