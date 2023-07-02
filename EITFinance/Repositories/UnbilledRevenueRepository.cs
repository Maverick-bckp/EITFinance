using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using EITFinance.Services;
using MailKit.Security;
using MimeKit;
using System.Linq;
using MailKit.Net.Smtp;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using EITFinance.Models;
using System.Data.Common;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Reflection;
using System.Text.RegularExpressions;

namespace EITFinance.Repositories
{
    public class UnbilledRevenueRepository : IUnbilledRevenueService
    {
        private IConfiguration _configuration;
        IHttpContextAccessor _httpContextAccessor;
        private readonly ILogger<SchedulerRepository> _logger;
        private readonly IWebHostEnvironment _HostEnvironment;
        static SqlConnection conn = null;

        public UnbilledRevenueRepository(IConfiguration Configuration, ILogger<SchedulerRepository> logger, IHttpContextAccessor httpContextAccessor, IWebHostEnvironment HostEnvironment)
        {
            _configuration = Configuration;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _HostEnvironment = HostEnvironment;
        }


        public bool uploadUnbilledRevenueFile(IFormFile file)
        {
            bool status = false;
            var timestamp = DateTime.Now.ToString("ddMMyyyyHHmm");
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-GB");

            HttpContextAccessor httpContextAccessor = new HttpContextAccessor();

            string loginID = httpContextAccessor.HttpContext.Session.GetString("username");

            SqlConnection con = new SqlConnection(_configuration.GetConnectionString("connEITFINDB"));
            SqlBulkCopy objbulk = new SqlBulkCopy(con);

            try
            {
                /*---- 0. Check File Extension ----*/
                if (file != null)
                {
                    var extension = Path.GetExtension(file.FileName);
                    if (extension.ToLower() != ".xlsx")
                    {
                        return status;
                    }
                }
                else
                {
                    return status;
                }

                /*---- 1. Delete 'collection_summary' Table---*/
                int deleteStatusUnbilledRevenue = deleteFromUnbilledRevenueTable(loginID);

                /*---- 2. Create DataTable ----*/
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("serial_no", typeof(string)));
                dt.Columns.Add(new DataColumn("resources_name", typeof(string)));
                dt.Columns.Add(new DataColumn("unique_id", typeof(string)));
                dt.Columns.Add(new DataColumn("av_series_no", typeof(string)));
                dt.Columns.Add(new DataColumn("emp_id", typeof(string)));
                dt.Columns.Add(new DataColumn("doj", typeof(string)));
                dt.Columns.Add(new DataColumn("billing_start_date", typeof(string)));
                dt.Columns.Add(new DataColumn("status", typeof(string)));
                dt.Columns.Add(new DataColumn("lwd", typeof(string)));
                dt.Columns.Add(new DataColumn("client", typeof(string)));
                dt.Columns.Add(new DataColumn("legal_entity_name", typeof(string)));
                dt.Columns.Add(new DataColumn("bu_type", typeof(string)));
                dt.Columns.Add(new DataColumn("nature_of_business", typeof(string)));
                dt.Columns.Add(new DataColumn("period", typeof(string)));
                dt.Columns.Add(new DataColumn("po_type", typeof(string)));
                dt.Columns.Add(new DataColumn("billing_spoc", typeof(string)));
                dt.Columns.Add(new DataColumn("currency", typeof(string)));
                dt.Columns.Add(new DataColumn("practice_area", typeof(string)));
                dt.Columns.Add(new DataColumn("ageing_bracket", typeof(string)));
                dt.Columns.Add(new DataColumn("source", typeof(string)));
                dt.Columns.Add(new DataColumn("po_no", typeof(string)));
                dt.Columns.Add(new DataColumn("location", typeof(string)));
                dt.Columns.Add(new DataColumn("zone", typeof(string)));
                dt.Columns.Add(new DataColumn("rates", typeof(string)));
                dt.Columns.Add(new DataColumn("unbilled_no_days", typeof(string)));
                dt.Columns.Add(new DataColumn("unbilled_quantity", typeof(string)));
                dt.Columns.Add(new DataColumn("unbilled_revenue", typeof(float)));
                dt.Columns.Add(new DataColumn("unbilled_category", typeof(string)));
                dt.Columns.Add(new DataColumn("unbilled_reason", typeof(string)));
                dt.Columns.Add(new DataColumn("account_manager", typeof(string)));
                dt.Columns.Add(new DataColumn("account_head", typeof(string)));
                dt.Columns.Add(new DataColumn("bu_head", typeof(string)));
                dt.Columns.Add(new DataColumn("ts_issue_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("po_issue_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("po_and_ts_issue_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("client_portal_issue_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("others_issue_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("bgv_issue_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("cost_paid_or_settled", typeof(string)));
                dt.Columns.Add(new DataColumn("cost_on_hold_or_settlement_due", typeof(string)));
                dt.Columns.Add(new DataColumn("status_as_on", typeof(string)));
                dt.Columns.Add(new DataColumn("remarks", typeof(string)));
                dt.Columns.Add(new DataColumn("file_uploader_id", typeof(string)));
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
                        for (int row = 2; row <= rowcount; row++)
                        {
                            DataRow dr = dt.NewRow();
                            dr["serial_no"] = worksheet.Cells[row, 1].Value == null ? "-" : worksheet.Cells[row, 1].Value.ToString().Trim();
                            dr["resources_name"] = worksheet.Cells[row, 2].Value == null ? "-" : worksheet.Cells[row, 2].Value.ToString().Trim();
                            dr["unique_id"] = worksheet.Cells[row, 3].Value == null ? "-" : worksheet.Cells[row, 3].Value.ToString().Trim();
                            dr["av_series_no"] = worksheet.Cells[row, 4].Value == null ? "-" : worksheet.Cells[row, 4].Value.ToString().Trim();
                            dr["emp_id"] = worksheet.Cells[row, 5].Value == null ? "-" : worksheet.Cells[row, 5].Value.ToString().Trim();
                            dr["doj"] = worksheet.Cells[row, 6].Value == null ? "-" : DateTime.FromOADate(double.Parse(worksheet.Cells[row, 6].Value.ToString().Trim())).ToString("dd-MMM-yyyy");
                            dr["billing_start_date"] = worksheet.Cells[row, 7].Value == null ? "-" : DateTime.FromOADate(double.Parse(worksheet.Cells[row, 7].Value.ToString().Trim())).ToString("dd-MMM-yyyy");
                            dr["status"] = worksheet.Cells[row, 8].Value == null ? "-" : worksheet.Cells[row, 8].Value.ToString().Trim();
                            dr["lwd"] = worksheet.Cells[row, 9].Value == null ? "-" : worksheet.Cells[row, 9].Value.ToString().Trim();
                            dr["client"] = worksheet.Cells[row, 10].Value == null ? "-" : worksheet.Cells[row, 10].Value.ToString().Trim();
                            dr["legal_entity_name"] = worksheet.Cells[row, 11].Value == null ? "-" : worksheet.Cells[row, 11].Value.ToString().Trim();
                            dr["bu_type"] = worksheet.Cells[row, 12].Value == null ? "-" : worksheet.Cells[row, 12].Value.ToString().Trim();
                            dr["nature_of_business"] = worksheet.Cells[row, 13].Value == null ? "-" : worksheet.Cells[row, 13].Value.ToString().Trim();
                            dr["period"] = worksheet.Cells[row, 14].Value == null ? "-" : worksheet.Cells[row, 14].Value.ToString().Trim();
                            dr["po_type"] = worksheet.Cells[row, 15].Value == null ? "-" : worksheet.Cells[row, 15].Value.ToString().Trim();
                            dr["billing_spoc"] = worksheet.Cells[row, 16].Value == null ? "-" : worksheet.Cells[row, 16].Value.ToString().Trim();
                            dr["currency"] = worksheet.Cells[row, 17].Value == null ? "-" : worksheet.Cells[row, 17].Value.ToString().Trim();
                            dr["practice_area"] = worksheet.Cells[row, 18].Value == null ? "-" : worksheet.Cells[row, 18].Value.ToString().Trim();
                            dr["ageing_bracket"] = worksheet.Cells[row, 19].Value == null ? "-" : worksheet.Cells[row, 19].Value.ToString().Trim();
                            dr["source"] = worksheet.Cells[row, 20].Value == null ? "-" : worksheet.Cells[row, 20].Value.ToString().Trim();
                            dr["po_no"] = worksheet.Cells[row, 21].Value == null ? "-" : worksheet.Cells[row, 21].Value.ToString().Trim();
                            dr["location"] = worksheet.Cells[row, 22].Value == null ? "-" : worksheet.Cells[row, 22].Value.ToString().Trim();
                            dr["zone"] = worksheet.Cells[row, 23].Value == null ? "-" : worksheet.Cells[row, 23].Value.ToString().Trim();
                            dr["rates"] = worksheet.Cells[row, 24].Value == null ? "-" : worksheet.Cells[row, 24].Value.ToString().Trim();
                            dr["unbilled_no_days"] = worksheet.Cells[row, 25].Value == null ? "-" : worksheet.Cells[row, 25].Value.ToString().Trim();
                            dr["unbilled_quantity"] = worksheet.Cells[row, 26].Value == null ? "-" : worksheet.Cells[row, 26].Value.ToString().Trim();
                            if (worksheet.Cells[row, 27].Value == null) { dr["unbilled_revenue"] = 0; } else { dr["unbilled_revenue"] = worksheet.Cells[row, 27].Value.ToString().Trim(); }
                            dr["unbilled_category"] = worksheet.Cells[row, 28].Value == null ? "-" : worksheet.Cells[row, 28].Value.ToString().Trim();
                            dr["unbilled_reason"] = worksheet.Cells[row, 29].Value == null ? "-" : worksheet.Cells[row, 29].Value.ToString().Trim();
                            dr["account_manager"] = worksheet.Cells[row, 30].Value == null ? "-" : worksheet.Cells[row, 30].Value.ToString().Trim();
                            dr["account_head"] = worksheet.Cells[row, 31].Value == null ? "-" : worksheet.Cells[row, 31].Value.ToString().Trim();
                            dr["bu_head"] = worksheet.Cells[row, 32].Value == null ? "-" : worksheet.Cells[row, 32].Value.ToString().Trim();
                            dr["ts_issue_amount"] = worksheet.Cells[row, 33].Value == null ? "-" : worksheet.Cells[row, 33].Value.ToString().Trim();
                            dr["po_issue_amount"] = worksheet.Cells[row, 34].Value == null ? "-" : worksheet.Cells[row, 34].Value.ToString().Trim();
                            dr["po_and_ts_issue_amount"] = worksheet.Cells[row, 35].Value == null ? "-" : worksheet.Cells[row, 35].Value.ToString().Trim();
                            dr["client_portal_issue_amount"] = worksheet.Cells[row, 36].Value == null ? "-" : worksheet.Cells[row, 36].Value.ToString().Trim();
                            dr["others_issue_amount"] = worksheet.Cells[row, 37].Value == null ? "-" : worksheet.Cells[row, 37].Value.ToString().Trim();
                            dr["bgv_issue_amount"] = worksheet.Cells[row, 38].Value == null ? "-" : worksheet.Cells[row, 38].Value.ToString().Trim();
                            dr["cost_paid_or_settled"] = worksheet.Cells[row, 39].Value == null ? "-" : worksheet.Cells[row, 39].Value.ToString().Trim();
                            dr["cost_on_hold_or_settlement_due"] = worksheet.Cells[row, 40].Value == null ? "-" : worksheet.Cells[row, 40].Value.ToString().Trim();
                            dr["status_as_on"] = worksheet.Cells[row, 41].Value == null ? "-" : worksheet.Cells[row, 41].Value.ToString().Trim();
                            dr["remarks"] = worksheet.Cells[row, 42].Value == null ? "-" : worksheet.Cells[row, 42].Value.ToString().Trim();
                            dr["file_uploader_id"] = loginID;
                            dr["timestamp"] = timestamp;

                            dt.Rows.Add(dr);
                        }
                    }
                }
                

                /*--- 5. Bulk Insert In 'billing_spoc' Table ---*/
                using (objbulk)
                {                    
                    objbulk.DestinationTableName = "unbilled_revenue";
                    objbulk.ColumnMappings.Add("serial_no", "serial_no");
                    objbulk.ColumnMappings.Add("resources_name", "resources_name");
                    objbulk.ColumnMappings.Add("unique_id", "unique_id");
                    objbulk.ColumnMappings.Add("av_series_no", "av_series_no");
                    objbulk.ColumnMappings.Add("emp_id", "emp_id");
                    objbulk.ColumnMappings.Add("doj", "doj");
                    objbulk.ColumnMappings.Add("billing_start_date", "billing_start_date");
                    objbulk.ColumnMappings.Add("status", "status");
                    objbulk.ColumnMappings.Add("lwd", "lwd");
                    objbulk.ColumnMappings.Add("client", "client");
                    objbulk.ColumnMappings.Add("legal_entity_name", "legal_entity_name");
                    objbulk.ColumnMappings.Add("bu_type", "bu_type");
                    objbulk.ColumnMappings.Add("nature_of_business", "nature_of_business");
                    objbulk.ColumnMappings.Add("period", "period");
                    objbulk.ColumnMappings.Add("po_type", "po_type");
                    objbulk.ColumnMappings.Add("billing_spoc", "billing_spoc");
                    objbulk.ColumnMappings.Add("currency", "currency");
                    objbulk.ColumnMappings.Add("practice_area", "practice_area");
                    objbulk.ColumnMappings.Add("ageing_bracket", "ageing_bracket");
                    objbulk.ColumnMappings.Add("source", "source");
                    objbulk.ColumnMappings.Add("po_no", "po_no");
                    objbulk.ColumnMappings.Add("location", "location");
                    objbulk.ColumnMappings.Add("zone", "zone");
                    objbulk.ColumnMappings.Add("rates", "rates");
                    objbulk.ColumnMappings.Add("unbilled_no_days", "unbilled_no_days");
                    objbulk.ColumnMappings.Add("unbilled_quantity", "unbilled_quantity");
                    objbulk.ColumnMappings.Add("unbilled_revenue", "unbilled_revenue");
                    objbulk.ColumnMappings.Add("unbilled_category", "unbilled_category");
                    objbulk.ColumnMappings.Add("unbilled_reason", "unbilled_reason");
                    objbulk.ColumnMappings.Add("account_manager", "account_manager");
                    objbulk.ColumnMappings.Add("account_head", "account_head");
                    objbulk.ColumnMappings.Add("bu_head", "bu_head");
                    objbulk.ColumnMappings.Add("ts_issue_amount", "ts_issue_amount");
                    objbulk.ColumnMappings.Add("po_issue_amount", "po_issue_amount");
                    objbulk.ColumnMappings.Add("po_and_ts_issue_amount", "po_and_ts_issue_amount");
                    objbulk.ColumnMappings.Add("client_portal_issue_amount", "client_portal_issue_amount");
                    objbulk.ColumnMappings.Add("others_issue_amount", "others_issue_amount");
                    objbulk.ColumnMappings.Add("bgv_issue_amount", "bgv_issue_amount");
                    objbulk.ColumnMappings.Add("cost_paid_or_settled", "cost_paid_or_settled");
                    objbulk.ColumnMappings.Add("cost_on_hold_or_settlement_due", "cost_on_hold_or_settlement_due");
                    objbulk.ColumnMappings.Add("status_as_on", "status_as_on");
                    objbulk.ColumnMappings.Add("remarks", "remarks");
                    objbulk.ColumnMappings.Add("file_uploader_id", "file_uploader_id");
                    objbulk.ColumnMappings.Add("timestamp", "timestamp");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(dt);
                    con.Close();
                }

                /*--- 6.Return With Success Status ---*/
                status = true;
                _logger.LogInformation("uploadUnbilledRevenueFile Method Success");
                return status;

            }
            catch (Exception ex)
            {
                string errorMessage = string.Empty;

                if (ex.Message.Contains("Received an invalid column length from the bcp client for colid"))
                {
                    // this method gives message with column name with length.  
                    errorMessage = GetBulkCopyColumnException(ex, objbulk);
                    // errorMessage contains "Column: "XYZ" contains data with a length greater than: 20", column, length  
                    Exception exInvlidColumn = new Exception(errorMessage, ex);

                }
                status = false;
                _logger.LogInformation("uploadUnbilledRevenueFile -------- " + ex.Message + ex.StackTrace);
                return status;
            }
        }

        protected string GetBulkCopyColumnException(Exception ex, SqlBulkCopy bulkcopy)

        {
            string message = string.Empty;
            if (ex.Message.Contains("Received an invalid column length from the bcp client for colid"))

            {
                string pattern = @"\d+";
                Match match = Regex.Match(ex.Message.ToString(), pattern);
                var index = Convert.ToInt32(match.Value) - 1;

                FieldInfo fi = typeof(SqlBulkCopy).GetField("_sortedColumnMappings", BindingFlags.NonPublic | BindingFlags.Instance);
                var sortedColumns = fi.GetValue(bulkcopy);
                var items = (Object[])sortedColumns.GetType().GetField("_items", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(sortedColumns);

                FieldInfo itemdata = items[index].GetType().GetField("_metadata", BindingFlags.NonPublic | BindingFlags.Instance);
                var metadata = itemdata.GetValue(items[index]);
                var column = metadata.GetType().GetField("column", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(metadata);
                var length = metadata.GetType().GetField("length", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance).GetValue(metadata);
                message = (String.Format("Column: {0} contains data with a length greater than: {1}", column, length));
            }
            return message;
        }

        public int deleteFromUnbilledRevenueTable(string loginID)
        {
            int status = 0;

            string sql_delete_billing_spoc_tbl = $"delete from unbilled_revenue where file_uploader_id = '{loginID}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_delete_billing_spoc_tbl;
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

        public bool sendUnbilledRevenueDetailsMail()
        {
            bool successFlag = false;
            var totalSum = 0f;
            var signature = "";
            string htmlBody = "";
            bool returnStatus = true;
            String attachmentPath = "";
            DataTable dt_client_list = new DataTable();
            DataTable dt_parent_category = new DataTable();
            DataTable dt_unbilled_revenue_details = new DataTable();
            DataTable dt_unbilled_revenue_summary = new DataTable();

            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string empID = _httpContextAccessor.HttpContext.Session.GetString("emp_id");

            try
            {
                /*----- Action : 1 ------*/
                dt_client_list = getUnbilledDataByActions(1, username);
                foreach (DataRow dr_client_list in dt_client_list.Rows)
                {
                    string client_name = dr_client_list[0].ToString();
                    /*--- Action : 2 ---*/
                    dt_parent_category = getUnbilledDataByActions(2, username, "", client_name);
                    foreach (DataRow dr_parent_category in dt_parent_category.Rows)
                    {
                        string parent_category = dr_parent_category[0].ToString();
                        string[] to_mail = dr_parent_category[1].ToString().Split(';');
                        string[] cc_mail = dr_parent_category[2].ToString().Split(';');

                        /*---- Clear Variables-----*/
                        successFlag = false;
                        htmlBody = string.Empty;


                        /*---- Declare Mail Subject -----*/
                        string subject = $@"Unbilled Revenue – {client_name} – {parent_category}";


                        /*---------- Phase 1 -------------*/
                        dt_unbilled_revenue_details = getUnbilledDataByActions(4, username, parent_category, client_name);
                        if (dt_unbilled_revenue_details.Rows.Count > 0)
                        {
                            attachmentPath = createExcelFile(dt_unbilled_revenue_details);

                            successFlag = true;
                        }




                        /*---------- Phase 2 -------------*/
                        dt_unbilled_revenue_summary = getUnbilledDataByActions(3, username, parent_category, client_name);
                        if (dt_unbilled_revenue_summary.Rows.Count > 0)
                        {
                            htmlBody = $@"<p style=""font-family:Calibri;font-size:15;"">Hi Team/Sir/Madam,</p></n><p style=""font-family:Calibri;font-size:15;"">Request your help in resolving the unbilled issue for {client_name} as per summary given below. The details are in the excel file attached herewith.</p></n><p style=""font-family:Calibri;font-size:15;"">Please revert on the same at the earliest or ignore this mail if already responded.</p>";

                            StringBuilder strHTMLBuilder = new StringBuilder();
                            strHTMLBuilder.Append("<table style=\"font-family:Calibri;font-size:12;border-collapse: collapse;width: 100%;\">");
                            strHTMLBuilder.Append("<tr >");
                            foreach (DataColumn myColumn in dt_unbilled_revenue_summary.Columns)
                            {
                                strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;\">");
                                strHTMLBuilder.Append(myColumn.ColumnName);
                                strHTMLBuilder.Append("</td>");
                            }
                            strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;\">Total</td>");
                            strHTMLBuilder.Append("</tr>");
                            foreach (DataRow myRow in dt_unbilled_revenue_summary.Rows)
                            {
                                strHTMLBuilder.Append("<tr >");
                                foreach (DataColumn myColumn in dt_unbilled_revenue_summary.Columns)
                                {
                                    strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;\">");
                                    strHTMLBuilder.Append(myRow[myColumn.ColumnName]);
                                    strHTMLBuilder.Append("</td>");
                                }

                                /*---- Summ of Issues Row Wise ----*/
                                totalSum = 0f;
                                for (int i = 1; i < dt_unbilled_revenue_summary.Columns.Count; i++)
                                {
                                    var vall = myRow[i].ToString();
                                    string issueColumnValue = myRow[i].ToString() == "" ? "0" : myRow[i].ToString();
                                    totalSum += float.Parse(issueColumnValue);
                                }
                                strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;\">");
                                strHTMLBuilder.Append(totalSum);
                                strHTMLBuilder.Append("</td>");


                                strHTMLBuilder.Append("</tr>");
                            }

                            /*---- Summ of Issues Row Wise ----*/

                            strHTMLBuilder.Append("<tr>");
                            strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;\">Total</td>");

                            string[] columnNames = dt_unbilled_revenue_summary.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
                            for (int i = 1; i < columnNames.Length; i++)
                            {
                                var columnSum = dt_unbilled_revenue_summary.AsEnumerable().Sum(x => x[columnNames[i]].ToString() == "" ? 0 : (double)x[columnNames[i]]);
                                var finalColumnSum = columnSum.ToString() == "0" ? "" : Math.Round(columnSum, 2).ToString();
                                strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;\">");
                                strHTMLBuilder.Append(finalColumnSum);
                                strHTMLBuilder.Append("</td>");
                            }
                            strHTMLBuilder.Append("<td style=\"border: 1px solid #dddddd;text-align: center;padding: 8px;\"></td>");
                            strHTMLBuilder.Append("</tr>");


                            strHTMLBuilder.Append("</table>");
                            string htmlTable = strHTMLBuilder.ToString();
                            htmlBody += htmlTable;

                            successFlag = true;
                        }

                        /*----- Add Sender Signature to Mail-----*/
                        string signaturePath = Path.Combine("wwwroot/Templates/Billing/", "" + empID + ".html");
                        if (File.Exists(signaturePath))
                        {
                            using (StreamReader streamReader = new StreamReader(signaturePath))
                            {
                                signature = streamReader.ReadToEnd();
                            }
                            htmlBody += signature;
                        }


                        /*------- Send Mail --------*/
                        try
                        {
                            if (successFlag == true)
                            {
                                sendMail(to_mail, cc_mail, htmlBody, subject, attachmentPath, username);
                                _logger.LogInformation("Unbilled Revenue Mail Send Finished For ---- " + client_name);
                                System.IO.File.Delete(attachmentPath);
                            }
                        }
                        catch (Exception exinner)
                        {
                            /*----- Log Exception Info Into File -----*/
                            _logger.LogInformation("SendMail Method -------- " + exinner.Message + exinner.StackTrace);
                        }
                    }

                    /*----- Log Info Into File -----*/
                    _logger.LogInformation("AutoMail Finished ---- " + client_name);
                }

                /*----- 0. Log Info Into File -----*/
                _logger.LogInformation("AutoMail Completed");
            }
            catch (Exception ex)
            {
                returnStatus = false;

                /*----- Log Exception Info Into File -----*/
                _logger.LogInformation(ex.StackTrace + "--------------" + ex.Message);
            }
            return returnStatus;
        }

        public dynamic getUnbilledDataByActions(int actionType = 0, string loginId = null, string parentCategory = "", string clientName = "")
        {
            string sp_distinct_cl = "get_unbilled_data_by_actions";
            DataTable dt_distinct_cl = new DataTable();

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_distinct_cl;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@Action", actionType);
                cmd.Parameters.AddWithValue("@login_id", loginId);
                cmd.Parameters.AddWithValue("@parent_category", parentCategory);
                cmd.Parameters.AddWithValue("@client_name", clientName);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlDataAdapter oda = new SqlDataAdapter(cmd);
                oda.Fill(dt_distinct_cl);
            }

            return dt_distinct_cl;
        }

        public void sendMail(string[] mailTo, string[] CCTo, string htmlBody, string subject, string attachmentPath, string loginID)
        {
            var host = _configuration.GetValue<string>("MailSettings:Host");
            var port = _configuration.GetValue<int>("MailSettings:Port");
            var username = _configuration.GetValue<string>("MailSettings:UserName");
            var password = _configuration.GetValue<string>("MailSettings:Password");
            var Sender = loginID;

            var body = new BodyBuilder();

            try
            {
                /*--- Sender ----*/
                var email = new MimeMessage();
                email.From.Add(MailboxAddress.Parse(Sender));

                /*--- Receiver ----*/
                foreach (string mailAddress in mailTo)
                {
                    if (!string.IsNullOrEmpty(mailAddress))
                    {
                        email.To.Add(MailboxAddress.Parse(mailAddress));
                    }
                }

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
                        body.Attachments.Add(attachment.FileName, memoryStream, ContentType.Parse("application/vnd.ms-excel"));
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

        private String createExcelFile(DataTable dataTable)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Worksheets.Add("UnbilledRevenueSummary");

                var excelWorksheet = excel.Workbook.Worksheets["UnbilledRevenueSummary"];

                var table = excelWorksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                string excelPath = Path.Combine("wwwroot/files/UnbilledRevenue/excel/", "" + "UnbilledRevenueDetails.xlsx");
                FileInfo excelFile = new FileInfo(excelPath);
                excel.SaveAs(excelFile);
                return excelPath;
            }
        }

        public bool uploadMaillingAddress(IFormFile file)
        {
            bool status = false;
            var timestamp = DateTime.Now.ToString("ddMMyyyyHHmm");

            HttpContextAccessor httpContextAccessor = new HttpContextAccessor();

            string loginID = httpContextAccessor.HttpContext.Session.GetString("username");

            try
            {
                if (file != null)
                {
                    var extension = Path.GetExtension(file.FileName);
                    if (extension.ToLower() != ".xlsx")
                    {
                        return status;
                    }
                }
                else
                {
                    return status;
                }

                /*---- 1. Delete 'unbilled mailling address' Table---*/
                int deleteMaillingAddressStaging = deleteFromUnbilledRevenueMaillingAddressStagingTable(loginID);

                /*---- 2. Create DataTable ----*/
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("pbrs_id", typeof(string)));
                dt.Columns.Add(new DataColumn("client_name", typeof(string)));
                dt.Columns.Add(new DataColumn("issue_category", typeof(string)));
                dt.Columns.Add(new DataColumn("to_mail", typeof(string)));
                dt.Columns.Add(new DataColumn("cc_mail", typeof(string)));
                dt.Columns.Add(new DataColumn("issue_parent_category", typeof(string)));
                dt.Columns.Add(new DataColumn("file_uploader_id", typeof(string)));

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
                        for (int row = 2; row <= rowcount; row++)
                        {
                            DataRow dr = dt.NewRow();
                            dr["pbrs_id"] = worksheet.Cells[row, 1].Value == null ? "-" : worksheet.Cells[row, 1].Value.ToString().Trim();
                            dr["client_name"] = worksheet.Cells[row, 2].Value == null ? "-" : worksheet.Cells[row, 2].Value.ToString().Trim();
                            dr["issue_category"] = worksheet.Cells[row, 3].Value == null ? "-" : worksheet.Cells[row, 3].Value.ToString().Trim();
                            dr["to_mail"] = worksheet.Cells[row, 4].Value == null ? "-" : worksheet.Cells[row, 4].Value.ToString().Trim();
                            dr["cc_mail"] = worksheet.Cells[row, 5].Value == null ? "-" : worksheet.Cells[row, 5].Value.ToString().Trim();
                            dr["issue_parent_category"] = worksheet.Cells[row, 6].Value == null ? "-" : worksheet.Cells[row, 6].Value.ToString().Trim();
                            dr["file_uploader_id"] = loginID;

                            dt.Rows.Add(dr);
                        }
                    }
                }

                /*--- 5. Bulk Insert In 'unbilled_mailing_address_staging' Table ---*/
                using (SqlConnection con = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    objbulk.DestinationTableName = "unbilled_mailing_address_staging";
                    objbulk.ColumnMappings.Add("pbrs_id", "pbrs_id");
                    objbulk.ColumnMappings.Add("client_name", "client_name");
                    objbulk.ColumnMappings.Add("issue_category", "issue_category");
                    objbulk.ColumnMappings.Add("to_mail", "to_mail");
                    objbulk.ColumnMappings.Add("cc_mail", "cc_mail");
                    objbulk.ColumnMappings.Add("issue_parent_category", "issue_parent_category");
                    objbulk.ColumnMappings.Add("file_uploader_id", "file_uploader_id");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(dt);
                    con.Close();
                }

                /*---- 6. Delete 'unbilled mailling address' Table---*/
                int deleteMaillingAddress = deleteFromUnbilledRevenueMaillingAddressTable(loginID);

                /*---- 7. Delete 'unbilled mailling address' Table---*/
                int deleteMaillingCategories = deleteFromUnbilledRevenueMaillingCategoriesTable(loginID);

                /*---- 8. Merge into 'Merge_to_Unbilled_Mailling_Categories' Table ---*/
                int mergeCategoriesStatus = mergeIntoUnbilledRevenueCategoriesTable();

                /*---- 9. Merge into 'Merge_to_Unbilled_Mailling_Address' Table ---*/
                int mergeAddressStatus = mergeIntoUnbilledRevenueAddressTable();

                /*--- 10. Return With Success Status ---*/
                status = true;
                _logger.LogInformation("uploadMaillingAddress Method Success");
            }
            catch (Exception ex)
            {
                _logger.LogInformation("uploadMaillingAddress -------- " + ex.Message + ex.StackTrace);
                status = false;
            }

            return status;
        }

        public int deleteFromUnbilledRevenueMaillingAddressStagingTable(string loginID)
        {
            int status = 0;

            string sql_delete_mail_addr_tbl = $"delete from unbilled_mailing_address_staging where file_uploader_id = '{loginID}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_delete_mail_addr_tbl;
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

        public int deleteFromUnbilledRevenueMaillingAddressTable(string loginID)
        {
            int status = 0;

            string sql_delete_mail_addr_tbl = $"delete from unbilled_mailing_address where file_uploader_id = '{loginID}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_delete_mail_addr_tbl;
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

        public int deleteFromUnbilledRevenueMaillingCategoriesTable(string loginID)
        {
            int status = 0;

            string sql_delete_mail_addr_tbl = $"delete from unbilled_mailing_categories where file_uploader_id = '{loginID}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_delete_mail_addr_tbl;
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

        public int mergeIntoUnbilledRevenueCategoriesTable()
        {
            int status = 0;

            string sp_merge_categories = "Merge_to_Unbilled_Mailling_Categories";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_merge_categories;
                cmd.CommandType = CommandType.StoredProcedure;
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                status = cmd.ExecuteNonQuery();
                conn.Close();
            }

            return status;
        }

        public int mergeIntoUnbilledRevenueAddressTable()
        {
            int status = 0;

            string sp_merge_address = "Merge_to_Unbilled_Mailling_Address";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_merge_address;
                cmd.CommandType = CommandType.StoredProcedure;
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                status = cmd.ExecuteNonQuery();
                conn.Close();
            }

            return status;
        }
    }
}
