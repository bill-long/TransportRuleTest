using System.Text.Json;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace Bilong
{
    public class TransportRuleTest
    {
        public static async Task Main(string[] args)
        {
            var config = JsonSerializer.Deserialize<EmailConfiguration>(File.ReadAllText(args[0]));

            if (config == null) return;

            Console.WriteLine($"Sending from {config.UserEmailAddress} to {config.RecipientEmailAddress} via {config.SmtpServer} with {(config.ClientId != null ? "OAuth2" : "Basic")} auth.");

            SaslMechanism saslMechanism;

            if (config.ClientId != null)
            {
                var authToken = await OAuthMethods.GetATokenForGraph(config.ClientId, config.Authority!, new[] {"https://outlook.office.com/SMTP.Send"});
                saslMechanism = new SaslMechanismOAuth2(authToken.Account.Username, authToken.AccessToken);
            }
            else
            {
                saslMechanism = new SaslMechanismLogin(config.UserEmailAddress, config.Password);
            }

            SendWeirdMime(
                new MailboxAddress(config.UserDisplayName, config.UserEmailAddress),
                new MailboxAddress(config.RecipientDisplayName, config.RecipientEmailAddress),
                config.SmtpServer,
                saslMechanism);
        }

        public static void SendWeirdMime(MailboxAddress from, MailboxAddress to, string? smtpServer, SaslMechanism saslMechanism)
        {
            var message = new MimeMessage();
            message.From.Add(from);
            message.To.Add(to);
            message.Subject = $"Transport rule test {DateTime.Now:o}";

            // The first body part will be plain text
            var bodyText = "Test";

            // The second body part will be the same text base64 encoded
            var memoryStream = new MemoryStream();
            var writer = new StreamWriter(memoryStream);
            writer.Write(bodyText);
            writer.Flush();
            memoryStream.Position = 0;

            var multipart = new MultipartAlternative();
            multipart.Add(new TextPart("plain") { Text = bodyText });
            multipart.Add(new MimePart("text", "plain")
            {
                Content = new MimeContent(memoryStream, ContentEncoding.Default),
                ContentTransferEncoding = ContentEncoding.Base64
            });

            message.Body = multipart;

            using var smtpClient = new SmtpClient();
            smtpClient.Connect(smtpServer, 587, MailKit.Security.SecureSocketOptions.StartTls);
            smtpClient.Authenticate(saslMechanism);
            smtpClient.Send(message);
        }
    }
}
