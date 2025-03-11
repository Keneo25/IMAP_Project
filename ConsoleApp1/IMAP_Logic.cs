using System.Runtime.InteropServices.Marshalling;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using Org.BouncyCastle.Tls;

namespace ConsoleApp1;

public class ImapLogic
{
    public void Run()
    {
        var cancellationTokenSource = new CancellationTokenSource();
        var cancellationToken = cancellationTokenSource.Token;
        using var client = new ImapClient();
        
        string host = Environment.GetEnvironmentVariable("IMAP_HOST") ?? throw new InvalidOperationException("Brak IMAP_HOST");
        string username = Environment.GetEnvironmentVariable("IMAP_USERNAME") ?? throw new InvalidOperationException("Brak IMAP_USERNAME");
        string password = Environment.GetEnvironmentVariable("IMAP_PASSWORD") ?? throw new InvalidOperationException("Brak IMAP_PASSWORD");
        bool success = int.TryParse(Environment.GetEnvironmentVariable("IMAP_PORT"), out int port);
        if (success == false)
        {
            throw new InvalidOperationException("Brak IMAP_PORT");
        }
        client.Connect(host, port, true);
        client.Authenticate(username,password);

        var inbox = client.Inbox;
        
        inbox.Open(FolderAccess.ReadWrite);
        var x = inbox.GetSubfolders().ToList();
        IMailFolder? destination = null;
        var exists = x.Any(p => p.Name == "OLD-RED");
        if (!exists)
        {
            destination = inbox.Create("OLD-RED", true);
        }
        else
        {
            destination = inbox.GetSubfolder("OLD-RED");
        }
        while (true)
        {
            CheckForNewMessages(inbox, destination);
            Thread.Sleep(TimeSpan.FromMinutes(1));
        }
    }
    private void CheckForNewMessages(IMailFolder inbox, IMailFolder destination)
    {
        var messages = inbox.Search(SearchQuery.SubjectContains("[RED]")).ToList();
        
        if (messages.Count == 0)
            return;
        
        ValidateMessages(inbox, destination, messages);
    }
    
    private void ValidateMessages(IMailFolder inbox, IMailFolder destination, List<UniqueId> messagesIds)
    {
        foreach (var messageId in messagesIds)
        {
            var message = inbox.GetMessage(messageId);
            if (message is null)
                continue;
                
            if (!message.Attachments.Any())
                continue;
                
            SaveAttachmentFromMessage(message.Attachments.ToList());
            inbox.MoveTo(messageId, destination);
        }
    }
    
    private void SaveAttachmentFromMessage(List<MimeEntity> attachments)
    {
        string path = "filepath";
        Directory.CreateDirectory(path);

        foreach (var attachment in attachments.OfType<MimePart>())
        {
            string name = attachment.ContentDisposition?.FileName ?? $"unknown_{Guid.NewGuid()}";
            string pathFile = Path.Combine(path, name);
            int counter = 1;

            while (File.Exists(pathFile))
            {
                string ex = Path.GetExtension(name);
                string filename = Path.GetFileNameWithoutExtension(name);
                pathFile = Path.Combine(path, $"{filename}({counter++}){ex}");
            }
            using var stream = File.Create(pathFile);
            attachment.Content.DecodeTo(stream); 
        }
    }
}