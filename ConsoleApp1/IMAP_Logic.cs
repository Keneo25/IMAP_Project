using System.Runtime.InteropServices.Marshalling;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using Org.BouncyCastle.Tls;

namespace ConsoleApp1;
public class ImapLogic
{
    private string Host { get; }
    private string Username { get; }
    private string Password { get; }
    private int Port { get; }
    
    public ImapLogic()
    {
        Host =  Environment.GetEnvironmentVariable("IMAP_HOST") ?? throw new InvalidOperationException("Brak IMAP_HOST");
        Username = Environment.GetEnvironmentVariable("IMAP_USERNAME") ?? throw new InvalidOperationException("Brak IMAP_USERNAME");
        Password = Environment.GetEnvironmentVariable("IMAP_PASSWORD") ?? throw new InvalidOperationException("Brak IMAP_PASSWORD");
        if (!int.TryParse(Environment.GetEnvironmentVariable("IMAP_PORT"), out int port))
        {
            throw new InvalidOperationException("Brak IMAP_PORT");
        }
        Port = port;
    }

    public void Connect(ImapClient client)
    {
        client.Connect(Host, Port, true);
        client.Authenticate(Username,Password);
    }

    public void CheckOrCreateFolder(ImapClient client)
    {
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
    
    public void Run()
    {
        using var client = new ImapClient();
        Connect(client);
        CheckOrCreateFolder(client);
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
        var path = "D:\\Pojekty\\Projek_Z_Trans\\fv";
        Directory.CreateDirectory(path);

        foreach (var attachment in attachments.OfType<MimePart>())
        {
            var name = attachment.ContentDisposition?.FileName ?? $"unknown_{Guid.NewGuid()}";
            var pathFile = Path.Combine(path, name);
            var counter = 1;

            while (File.Exists(pathFile))
            {
                var ex = Path.GetExtension(name);
                var filename = Path.GetFileNameWithoutExtension(name);
                pathFile = Path.Combine(path, $"{filename}({counter++}){ex}");
            }
            using var stream = File.Create(pathFile);
            attachment.Content.DecodeTo(stream); 
        }
    }
}