using System.CommandLine;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace Bilong
{
    public class TransportRuleTest
    {
        public static async Task<int> Main(string[] args)
        {
            var rootCommand = new RootCommand("Tool to send multipart/alternative SMTP mail to Exchange On-Prem or Online for testing.");
            var fromArgument = new Argument<string>(
                name: "from",
                description: "From address. A display name can be included by providing the value as " +
                    "\"John Doe <jdoe@contoso.com>\". This is also the UPN used when authenticating."
            );
            var toArgument = new Argument<string>(
                name: "to",
                description: "To address. This can include a display name just as in the From."
            );
            var smtpServerArgument = new Argument<string>(
                name: "server",
                description: "The SMTP server to use."
            );
            var portNumberOption = new Option<int>(
                name: "--port",
                getDefaultValue: () => 587,
                description: "The port number to use."
            );
            var passwordOption = new Option<string>(
                name: "--password",
                description: "If provided, AUTH LOGIN is used (basic)."
            );
            var clientIdOption = new Option<string>(
                name: "--clientId",
                description: "If provided, XOAUTH2 is used."
            );
            var tenantIdOption = new Option<string>(
                name: "--tenantId",
                description: "The tenant ID. Must be provided if clientId is provided."
            );

            rootCommand.AddArgument(fromArgument);
            rootCommand.AddArgument(toArgument);
            rootCommand.AddArgument(smtpServerArgument);
            rootCommand.AddOption(portNumberOption);
            rootCommand.AddOption(passwordOption);
            rootCommand.AddOption(clientIdOption);
            rootCommand.AddOption(tenantIdOption);

            rootCommand.SetHandler(async (fromValue, toValue, smtpServerValue, portNumberValue, passwordValue, clientIdValue, tenantIdValue) =>
            {
                SaslMechanism saslMechanism;
                var from = MailboxAddress.Parse(fromValue);
                var to = MailboxAddress.Parse(toValue);

                if (string.IsNullOrEmpty(passwordValue))
                {
                    if (string.IsNullOrEmpty(clientIdValue))
                    {
                        Console.WriteLine($"Either {clientIdOption.Name} or {passwordOption.Name} must be provided.");
                        return;
                    }
                    else if (string.IsNullOrEmpty(tenantIdValue))
                    {
                        Console.WriteLine($"{tenantIdOption.Name} must be provided with {clientIdOption.Name}.");
                        return;
                    }
                }

                Console.WriteLine($"Sending from {from.Address} to {to.Address} via {smtpServerValue}:{portNumberValue} with {(string.IsNullOrEmpty(clientIdValue) ? "Basic" : "OAuth2")} auth.");

                if (string.IsNullOrEmpty(clientIdValue))
                {
                    saslMechanism = new SaslMechanismLogin(from.Address, passwordValue);
                }
                else
                {
                    var authToken = await OAuthMethods.GetATokenForGraph(clientIdValue, $"https://login.microsoftonline.com/{tenantIdValue}", new[] { "https://outlook.office.com/SMTP.Send" });
                    saslMechanism = new SaslMechanismOAuth2(authToken.Account.Username, authToken.AccessToken);
                }

                SendMultipartAlternative(from, to, smtpServerValue, portNumberValue, saslMechanism);
            },
            fromArgument, toArgument, smtpServerArgument, portNumberOption, passwordOption, clientIdOption, tenantIdOption);

            return await rootCommand.InvokeAsync(args);
        }

        public static void SendMultipartAlternative(MailboxAddress from, MailboxAddress to, string? smtpServer, int portNumber, SaslMechanism saslMechanism)
        {
            var message = new MimeMessage();
            message.From.Add(from);
            message.To.Add(to);
            message.Subject = $"Transport rule test {DateTime.Now:o}";

            var multipart = new MultipartAlternative();
            multipart.Add(new TextPart("html") { Text = "<!DOCTYPE html><html lang=\"en\"><body>HTML Body</body></html>" });
            multipart.Add(new TextPart("foo") { Text = "Bad body part" });

            message.Body = multipart;

            using var smtpClient = new SmtpClient();
            smtpClient.Connect(smtpServer, portNumber, MailKit.Security.SecureSocketOptions.StartTls);
            smtpClient.Authenticate(saslMechanism);
            smtpClient.Send(message);
        }
    }
}
