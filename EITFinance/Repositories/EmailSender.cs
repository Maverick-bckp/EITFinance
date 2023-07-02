using EITFinance.Models;
using EITFinance.Services;
using MailKit.Security;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MimeKit;
using System;

namespace EITFinance.Repositories
{
    public class EmailSender : IEmailSender
    {
        private readonly IConfiguration _configuration;
        private readonly ILogger<EmailSender> _logger;
        public EmailSender(IConfiguration configuration, ILogger<EmailSender> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public void SendEmail(Message message)
        {
            var emailMessage = CreateEmailMessage(message);
            Send(emailMessage);
        }
        private MimeMessage CreateEmailMessage(Message message)
        {
            var emailMessage = new MimeMessage();
            emailMessage.From.Add(new MailboxAddress("Experis IT", "ayan.roy@in.experis.com"));
            emailMessage.To.AddRange(message.To);

            if(message.Cc.Count>0)
            emailMessage.Cc.AddRange(message.Cc);

            emailMessage.Subject = message.Subject;
            emailMessage.Body = new TextPart(MimeKit.Text.TextFormat.Html) { Text = string.Format("<h2 style='color:red;'>{0}</h2>", message.Content) };

            return emailMessage;
        }

        private void Send(MimeMessage mailMessage)
        {
            using (var client = new MailKit.Net.Smtp.SmtpClient())
            {
                try
                {
                    var host = _configuration.GetValue<string>("MailSettings:Host");
                    var port = _configuration.GetValue<int>("MailSettings:Port");
                    var username = _configuration.GetValue<string>("MailSettings:UserName");
                    var password = _configuration.GetValue<string>("MailSettings:Password");
                    var Sender = _configuration.GetValue<string>("MailSettings:Sender");

                    client.Connect(host, port, SecureSocketOptions.Auto);
                    client.Send(mailMessage);                    
                }
                catch(Exception ex)
                {
                    _logger.LogError(ex.Message);
                }
            }
        }
    }
}
