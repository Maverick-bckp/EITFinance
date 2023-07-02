using Microsoft.Extensions.Configuration;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using System.Linq;
using EITFinance.Models;
using EITFinance.Services;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text;
using System;
using System.Globalization;

namespace EITFinance.Repositories
{
    public class SMTPRepository : ISMTPService
    {
        IConfiguration _configuration;
        public SMTPRepository(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public void sendMail(CollectionSummaryMailData mailData)
        {
            var signature = "<h4>Thanks & Regards,</h4></n><h4>Experis AR Team</h4>";
            int sl_no_counter = 1;
            var host = _configuration.GetValue<string>("MailSettings:Host");
            var port = _configuration.GetValue<int>("MailSettings:Port");
            var username = _configuration.GetValue<string>("MailSettings:UserName");
            var password = _configuration.GetValue<string>("MailSettings:Password");
            var Sender = _configuration.GetValue<string>("MailSettings:Sender");

            /*--- Sender ----*/
            var email = new MimeMessage();
            email.From.Add(MailboxAddress.Parse(Sender));

            /*--- Receiver ----*/
            foreach (string mailAddress in mailData.mailTo)
                email.To.Add(MailboxAddress.Parse(mailAddress));

            /*--- CC ----*/
            /*--- Check if a CC address was supplied in the request ----*/
            if (mailData.CCTo != null)
            {
                foreach (string mailAddress in mailData.CCTo.Where(x => !string.IsNullOrWhiteSpace(x.ToString())))
                    email.Cc.Add(MailboxAddress.Parse(mailAddress.Trim()));
            }

            /*--- Add Content to Mime Message ----*/
            var body = new BodyBuilder();
            email.Subject = $"Remittance Advice Required - {mailData.clientName}";
            string htmlBody = @"<p style=""font-family:Calibri;font-size:15;"">Dear Sir/Madam,</p></n><p style=""font-family:Calibri;font-size:15;"">Kindly help us with the invoice-wise adjustment details (Remittance Advice) for the below payments urgently.</p></br>";

            htmlBody += @"<table style=""font-family:Calibri;font-size:12;border-collapse: collapse;width: 100%;"">";
            htmlBody += @"<tr>
		                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">SL No.</th>
		                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Online Transfer/Cheque Details</th>
		                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Payment Received Date</th>
		                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Amount</th>
		                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Currency</th>
	                      </tr>";
            foreach (JObject collData in mailData.collectionData)
            {
                htmlBody += $@"<tr>
		                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{sl_no_counter}</td>
		                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{(collData["currency"].ToString() == "INR" ? collData["cheque_details"] : "-")}</td>
		                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{collData["payment_received_date"]}</td>
                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{String.Format(new CultureInfo( "en-IN", false ), "{0:n}", Convert.ToDouble(collData["amount"]))}</td>
		                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{collData["currency"]}</td>
	                        </tr>";

                sl_no_counter++;
            }

            htmlBody += @"</table>";
            htmlBody += @"<br><br>";

            using (StreamReader streamReader = new StreamReader(Path.Combine("wwwroot/Templates", "EIT_Signature.html")))
            {
                signature = streamReader.ReadToEnd();
            }
            htmlBody += signature;


            body.HtmlBody = string.Format(htmlBody) ;
            email.Body = body.ToMessageBody();

            /*--- Connect SMTP server and send mail ----*/
            using var smtp = new SmtpClient();
            smtp.Connect(host, port, SecureSocketOptions.Auto);
            smtp.Send(email);
            smtp.Disconnect(true);

        }
    }
}
