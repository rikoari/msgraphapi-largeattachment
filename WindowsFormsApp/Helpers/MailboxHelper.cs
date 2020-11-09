using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Graph;
using System.Linq;
using System.IO;

namespace WindowsFormsApp
{
    public class MailboxHelper
    {
        private GraphServiceClient _graphClient;
        private HttpClient _httpClient;
        
        public MailboxHelper(GraphServiceClient graphClient)
        {
            if (null == graphClient) throw new ArgumentNullException(nameof(graphClient));
                _graphClient = graphClient;
        }

        public MailboxHelper(HttpClient httpClient)
        {
            if (null == httpClient) throw new ArgumentNullException(nameof(httpClient));
                _httpClient = httpClient;
        }

        public MailboxHelper(GraphServiceClient graphClient, HttpClient httpClient)
        {
            if (null == graphClient) throw new ArgumentNullException(nameof(graphClient));
                _graphClient = graphClient;
            if (null == httpClient) throw new ArgumentNullException(nameof(httpClient));
                _httpClient = httpClient;
        }

        // Create message
        public async Task<Message> CreateMessage(Message message, string userId)
        {
            return await _graphClient
                .Users[userId]
                .Messages
                .Request()
                .AddAsync(message);
        }

        // Create upload session
        public async Task<UploadSession> CreateUploadSession(Message message, string userId, AttachmentItem attachmentItem)
        {
            return await _graphClient
                .Users[userId]
                .Messages[message.Id]
                .Attachments
                .CreateUploadSession(attachmentItem)
                .Request()
                .PostAsync();
        }

        // Upload attachments
        public async Task<UploadResult<AttachmentItem>> UploadAttachment(UploadSession uploadSession, FileStream attachmentStream)
        {
            int fileSlice = 320 * 1024;

            var fileUploadTask = new LargeFileUploadTask<AttachmentItem>(uploadSession, attachmentStream, fileSlice);

            IProgress<long> progress = new Progress<long>(progress => {
                Console.WriteLine($"Uploaded {progress} bytes of {attachmentStream.Length} bytes");
            });

            // Upload the file
            return await fileUploadTask.UploadAsync(progress);
        }

        // Send message
        public async void SendMessage(Message message, string userId)
        {
            await _graphClient
                .Users[userId]
                .Messages[message.Id]
                .Send()
                .Request()
                .PostAsync();
        }


    }
}