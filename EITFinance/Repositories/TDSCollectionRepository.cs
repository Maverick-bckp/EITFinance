using EITFinance.Services;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MimeKit;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace EITFinance.Repositories
{
    public class TDSCollectionRepository : ITDSCollectionService
    {
        private IConfiguration _configuration;
        IHttpContextAccessor _httpContextAccessor;
        private readonly ILogger<SchedulerRepository> _logger;
        private readonly IWebHostEnvironment _HostEnvironment;
        static SqlConnection conn = null;
        public TDSCollectionRepository(IConfiguration Configuration, ILogger<SchedulerRepository> logger, IHttpContextAccessor httpContextAccessor, IWebHostEnvironment HostEnvironment)
        {
            _configuration = Configuration;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _HostEnvironment = HostEnvironment;
        }

        public bool UploadTDSCollectionFile(IFormFile file)
        {
            bool status = false;
            var timestamp = DateTime.Now.ToString("ddMMyyyyHHmm");
            string loginID = _httpContextAccessor.HttpContext.Session.GetString("username");

            try
            {
                /*---- 0. Check File Extension ----*/
                if (file != null)
                {
                    var extension = Path.GetExtension(file.FileName);
                    if (extension.ToLower() != ".xlsx")
                    {
                        status = false;
                        return status;
                    }
                }
                else
                {
                    status = false;
                    return status;
                }

                /*---- 1. Truncate 'collection_summary' Table---*/
                int deleteStatusBillingSPOC = deleteFromTDSCollectionLogTable(loginID);

                /*---- 2. Create DataTable ----*/
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("name_of_deductor", typeof(string)));
                dt.Columns.Add(new DataColumn("tan_of_deductor", typeof(string)));
                dt.Columns.Add(new DataColumn("experis_pan_number", typeof(string)));
                dt.Columns.Add(new DataColumn("revenue_as_per_26as", typeof(string)));
                dt.Columns.Add(new DataColumn("tds_as_per_26as", typeof(string)));
                dt.Columns.Add(new DataColumn("tds_as_per_experis_books", typeof(string)));
                dt.Columns.Add(new DataColumn("tds_short_amount_as_per_26as", typeof(string)));
                dt.Columns.Add(new DataColumn("tds_excess_amount_as_per_26as", typeof(string)));
                dt.Columns.Add(new DataColumn("net_short_excess", typeof(string)));
                dt.Columns.Add(new DataColumn("financial_year", typeof(string)));
                dt.Columns.Add(new DataColumn("q1", typeof(string)));
                dt.Columns.Add(new DataColumn("q2", typeof(string)));
                dt.Columns.Add(new DataColumn("q3", typeof(string)));
                dt.Columns.Add(new DataColumn("q4", typeof(string)));
                dt.Columns.Add(new DataColumn("total_tds_certificates_received", typeof(string)));
                dt.Columns.Add(new DataColumn("total_pending_to_collect", typeof(string)));
                dt.Columns.Add(new DataColumn("to_mail", typeof(string)));
                dt.Columns.Add(new DataColumn("cc_mail", typeof(string)));
                dt.Columns.Add(new DataColumn("file_uploader_id", typeof(string)));
                dt.Columns.Add(new DataColumn("sent_status", typeof(string)));
                dt.Columns.Add(new DataColumn("timestamp", typeof(string)));

                /*--- 3. Read Data From Excel ---*/
                /*--- 4. Insert Values In DataTable After Reading Excel Data ---*/
                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var package = new ExcelPackage(stream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        var rowcount = worksheet.Dimension.Rows;
                        for (int row = 4; row <= rowcount; row++)
                        {
                            var client_name = worksheet.Cells[row, 3].Value;

                            if (client_name == null)
                            {
                                continue;
                            }
                            if (client_name.ToString() == "" || client_name.ToString() == "-")
                            {
                                continue;
                            }
                            if (client_name.ToString().ToLower().Equals("grand total"))
                            {
                                break;
                            }

                            DataRow dr = dt.NewRow();
                            dr["name_of_deductor"] = worksheet.Cells[row, 3].Value == null ? "-" : worksheet.Cells[row, 3].Value.ToString().Trim();
                            dr["tan_of_deductor"] = worksheet.Cells[row, 4].Value == null ? "-" : worksheet.Cells[row, 4].Value.ToString().Trim();
                            dr["experis_pan_number"] = worksheet.Cells[row, 5].Value == null ? "-" : worksheet.Cells[row, 5].Value.ToString().Trim();
                            dr["revenue_as_per_26as"] = worksheet.Cells[row, 6].Value == null ? "-" : worksheet.Cells[row, 6].Value.ToString().Trim();
                            dr["tds_as_per_26as"] = worksheet.Cells[row, 7].Value == null ? "-" : worksheet.Cells[row, 7].Value.ToString().Trim();
                            dr["tds_as_per_experis_books"] = worksheet.Cells[row, 8].Value == null ? "-" : worksheet.Cells[row, 8].Value.ToString().Trim();
                            dr["tds_short_amount_as_per_26as"] = worksheet.Cells[row, 9].Value == null ? "-" : worksheet.Cells[row, 9].Value.ToString().Trim();
                            dr["tds_excess_amount_as_per_26as"] = worksheet.Cells[row, 10].Value == null ? "-" : worksheet.Cells[row, 10].Value.ToString().Trim();
                            dr["net_short_excess"] = worksheet.Cells[row, 11].Value == null ? "-" : worksheet.Cells[row, 11].Value.ToString().Trim();
                            dr["financial_year"] = worksheet.Cells[row, 12].Value == null ? "-" : worksheet.Cells[row, 12].Value.ToString().Trim();
                            dr["q1"] = worksheet.Cells[row, 13].Value.ToString() == "-" ? "0" : worksheet.Cells[row, 13].Value.ToString().Trim();
                            dr["q2"] = worksheet.Cells[row, 14].Value.ToString() == "-" ? "0" : worksheet.Cells[row, 14].Value.ToString().Trim();
                            dr["q3"] = worksheet.Cells[row, 15].Value.ToString() == "-" ? "0" : worksheet.Cells[row, 15].Value.ToString().Trim();
                            dr["q4"] = worksheet.Cells[row, 16].Value.ToString() == "-" ? "0" : worksheet.Cells[row, 16].Value.ToString().Trim();
                            dr["total_tds_certificates_received"] = worksheet.Cells[row, 17].Value == null ? "-" : worksheet.Cells[row, 17].Value.ToString().Trim();
                            dr["total_pending_to_collect"] = worksheet.Cells[row, 18].Value == null ? "-" : worksheet.Cells[row, 18].Value.ToString().Trim();
                            dr["to_mail"] = worksheet.Cells[row, 20].Value == null ? "-" : worksheet.Cells[row, 20].Value.ToString().Trim();
                            dr["cc_mail"] = worksheet.Cells[row, 21].Value == null ? "-" : worksheet.Cells[row, 21].Value.ToString().Trim();

                            dr["file_uploader_id"] = loginID;
                            dr["timestamp"] = timestamp;

                            dt.Rows.Add(dr);
                        }
                    }
                }

                /*--- 5. Bulk Insert In 'billing_spoc' Table ---*/
                using (SqlConnection con = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    objbulk.DestinationTableName = "tds_collection_log";

                    objbulk.ColumnMappings.Add("name_of_deductor", "name_of_deductor");
                    objbulk.ColumnMappings.Add("tan_of_deductor", "tan_of_deductor");
                    objbulk.ColumnMappings.Add("experis_pan_number", "experis_pan_number");
                    objbulk.ColumnMappings.Add("revenue_as_per_26as", "revenue_as_per_26as");
                    objbulk.ColumnMappings.Add("tds_as_per_26as", "tds_as_per_26as");
                    objbulk.ColumnMappings.Add("tds_as_per_experis_books", "tds_as_per_experis_books");
                    objbulk.ColumnMappings.Add("tds_short_amount_as_per_26as", "tds_short_amount_as_per_26as");
                    objbulk.ColumnMappings.Add("tds_excess_amount_as_per_26as", "tds_excess_amount_as_per_26as");
                    objbulk.ColumnMappings.Add("net_short_excess", "net_short_excess");
                    objbulk.ColumnMappings.Add("financial_year", "financial_year");
                    objbulk.ColumnMappings.Add("q1", "q1");
                    objbulk.ColumnMappings.Add("q2", "q2");
                    objbulk.ColumnMappings.Add("q3", "q3");
                    objbulk.ColumnMappings.Add("q4", "q4");
                    objbulk.ColumnMappings.Add("total_tds_certificates_received", "total_tds_certificates_received");
                    objbulk.ColumnMappings.Add("total_pending_to_collect", "total_pending_to_collect");
                    objbulk.ColumnMappings.Add("to_mail", "to_mail");
                    objbulk.ColumnMappings.Add("cc_mail", "cc_mail");
                    objbulk.ColumnMappings.Add("file_uploader_id", "file_uploader_id");
                    objbulk.ColumnMappings.Add("timestamp", "timestamp");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(dt);
                    con.Close();
                }

                /*--- 6.Return With Success Status ---*/
                status = true;
                _logger.LogInformation("UploadTDSCollectionFile Method Success");
                return status;
            }
            catch (Exception ex)
            {
                status = false;
                _logger.LogInformation("UploadTDSCollectionFile -------- " + ex.Message + ex.StackTrace);
                return status;
            }
        }

        public int deleteFromTDSCollectionLogTable(string loginID)
        {
            int status = 0;

            string sql_delete_tds_col_log_tbl = $"delete from tds_collection_log where file_uploader_id = '{loginID}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_delete_tds_col_log_tbl;
                cmd.CommandType = CommandType.Text;
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                status = cmd.ExecuteNonQuery();
                conn.Close();
            }

            return status;
        }

        public dynamic getTDSCollectionLog(string loginId)
        {
            string sp_tds_collection_log = "get_tds_collection_log";
            DataTable dt_tds_collection_log = new DataTable();

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_tds_collection_log;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@login_id", loginId);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlDataAdapter oda = new SqlDataAdapter(cmd);
                oda.Fill(dt_tds_collection_log);
            }

            return dt_tds_collection_log;
        }

        public int updateMailSendStatus(string clientName)
        {
            int status = 0;

            string sql_update_mail_send_status = $"update tds_collection_log set sent_status='true' where name_of_deductor='{clientName}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_update_mail_send_status;
                cmd.CommandType = CommandType.Text;
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                status = cmd.ExecuteNonQuery();
                conn.Close();
            }

            return status;
        }

        public void sendMail(string[] mailTo, string[] CCTo, string htmlBody, string subject, string attachmentPath)
        {
            var host = _configuration.GetValue<string>("MailSettings:Host");
            var port = _configuration.GetValue<int>("MailSettings:Port");
            var username = _configuration.GetValue<string>("MailSettings:UserName");
            var password = _configuration.GetValue<string>("MailSettings:Password");
            var Sender = _configuration.GetValue<string>("MailSettings:Sender");

            var body = new BodyBuilder();

            try
            {
                /*--- Sender ----*/
                var email = new MimeMessage();
                email.From.Add(MailboxAddress.Parse(Sender));

                /*--- Receiver ----*/
                foreach (string mailAddress in mailTo)
                    email.To.Add(MailboxAddress.Parse(mailAddress));

                /*--- CC ----*/
                /*--- Check if a CC address was supplied in the request ----*/
                if (CCTo != null)
                {
                    foreach (string mailAddress in CCTo.Where(x => !string.IsNullOrWhiteSpace(x.ToString())))
                        email.Cc.Add(MailboxAddress.Parse(mailAddress.Trim()));
                }


                /*--- Create a new memory stream and attach attachment to mail body ---*/
                if (File.Exists(attachmentPath))
                {
                    using (MemoryStream memoryStream = new MemoryStream(File.ReadAllBytes(attachmentPath).ToArray()))
                    {
                        /*--- Copy the attachment to the stream --*/
                        var attachment = new FormFile(memoryStream, 0, memoryStream.Length, "streamFile", attachmentPath.Split(@"/").Last());

                        /*--Add the attachment from the byte array--*/
                        body.Attachments.Add(attachment.FileName, memoryStream, ContentType.Parse("application/pdf"));
                    }
                }


                /*--- Create Body----*/
                email.Subject = subject;
                body.HtmlBody = string.Format(htmlBody);
                email.Body = body.ToMessageBody();

                /*--- Connect SMTP server and send mail ----*/
                using var smtp = new SmtpClient();
                smtp.Connect(host, port, SecureSocketOptions.Auto);
                smtp.Send(email);
                smtp.Disconnect(true);

            }
            catch (Exception ex)
            {
                _logger.LogInformation("sendMail ---- " + ex.Message + ex.StackTrace);
            }
        }

        public bool sendTDSCollectionMail()
        {
            bool returnStatus = false;
            string clientName = "";
            string htmlBody = "";
            var signature = "";
            string[] toMail = { };
            string[] ccMail = { };
            string mailSubject = "";
            DataTable dt_tdsCollectionLog = new DataTable();
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string empID = _httpContextAccessor.HttpContext.Session.GetString("emp_id");


            /*----- 0. Log Info Into File -----*/
            _logger.LogInformation("TDS Collection Log Mail Sent Started");

            try
            {
                /*---------- Get All Data Of Billing SPOC ----------*/
                dt_tdsCollectionLog = getTDSCollectionLog(username);

                if (dt_tdsCollectionLog.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt_tdsCollectionLog.Rows)
                    {
                        clientName = dr["name_of_deductor"].ToString();
                        if (string.IsNullOrEmpty(clientName) || clientName == "-")
                        {
                            continue;
                        }

                        /*--- Create Html Body---*/
                        htmlBody = @"<p style=""font-family:Calibri;font-size:15;"">Hi Sir/Madam,</p></n><p style=""font-family:Calibri;font-size:15;"">Please share the TDS certificate for the below period along with the invoice-wise details for TDS reconciliation purposes.</br></br> Also please share the invoice wise details even if the TDS certificate has been shared for our reconciliation. </p></br>";
                        htmlBody += @"<table style=""font-family:Calibri;font-size:12;border-collapse: collapse;width: 100%;"">";
                        htmlBody += @"<tr>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Name of Deductor</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">TAN of Deductor</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Experis PAN Number</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Revenue as per 26AS</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">TDS As per 26AS</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">TDS As per Experis Books</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">TDS Short amount as per 26AS</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">TDS Excess amount as per 26AS</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">NET SHORT/EXCESS</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Financial Year</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;""> Q1</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;""> Q2</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;""> Q3</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;""> Q4</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">  Total TDS certificates received (Value in Rs)</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">   Total Pending to collect (Value in Rs)</th>
	                              </tr>";
                        htmlBody += $@"<tr>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["name_of_deductor"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["tan_of_deductor"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["experis_pan_number"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["revenue_as_per_26as"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["tds_as_per_26as"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{float.Parse(dr["tds_as_per_experis_books"].ToString()).ToString("0.00")}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{float.Parse(dr["tds_short_amount_as_per_26as"].ToString()).ToString("0.00")}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{float.Parse(dr["tds_excess_amount_as_per_26as"].ToString()).ToString("0.00")}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{float.Parse(dr["net_short_excess"].ToString()).ToString("0.00")}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["financial_year"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["q1"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["q2"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["q3"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["q4"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["total_tds_certificates_received"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dr["total_pending_to_collect"].ToString()}</td>
	                                </tr>";
                        htmlBody += @"</table>";
                        htmlBody += @"<br><br>";

                        /*---- Add Signature ----*/
                        using (StreamReader streamReader = new StreamReader(Path.Combine("wwwroot/Templates", "EIT_Signature.html")))
                        {
                            signature = streamReader.ReadToEnd();
                        }
                        htmlBody += signature;

                        /*---- Set To and CC ----*/
                        toMail = dr["to_mail"].ToString().Split(';');
                        ccMail = dr["cc_mail"].ToString().Split(';');

                        /*---- Mail Subject ----*/
                        mailSubject = $"Required TDS Certificate – {clientName}";

                        /*--- Send Mail ----*/
                        try
                        {
                            sendMail(toMail, ccMail, htmlBody, mailSubject, null);
                            _logger.LogInformation("Billing Mail Finished For ---- " + clientName);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogInformation("TDS Collection Mail Send Error For ---- " + clientName + ex.Message + ex.StackTrace);
                        }


                        /*---- Update Invoice No Send Status ----*/
                        /*--- So that Doesn't mail Duplicate Again---*/
                        int sentUpdateStatus = updateMailSendStatus(clientName);
                        _logger.LogInformation("TDS Collection Updated Sent Status Of ---- " + clientName);

                        /*----- 0. Log Info Into File -----*/
                        _logger.LogInformation("TDS Collection Mail Sent Complete For All");
                    }


                    returnStatus = true;
                }
                else
                {
                    /*----- 0. Log Info Into File -----*/
                    _logger.LogInformation("No Data Found To Mail !!");

                    returnStatus = false;
                    return returnStatus;
                }
            }
            catch (Exception ex)
            {
                /*----- 0. Log Info Into File -----*/
                _logger.LogInformation("sendTDSCollectionMail Method ---- " + ex.Message + ex.StackTrace);
                returnStatus = false;
            }

            return returnStatus;
        }
    }
}
