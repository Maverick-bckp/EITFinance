using EITFinance.Services;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using MimeKit;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace EITFinance.Repositories
{
    public class BillingRepository : IBillingService
    {
        private IConfiguration _configuration;
        IHttpContextAccessor _httpContextAccessor;
        private readonly ILogger<SchedulerRepository> _logger;
        private readonly IWebHostEnvironment _HostEnvironment;
        static SqlConnection conn = null;
        public BillingRepository(IConfiguration Configuration, ILogger<SchedulerRepository> logger, IHttpContextAccessor httpContextAccessor, IWebHostEnvironment HostEnvironment)
        {
            _configuration = Configuration;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _HostEnvironment=HostEnvironment;
        }

        public bool uploadBillingFile(IFormFile file)
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
                int deleteStatusBillingSPOC = deleteBillingSPOCTable(loginID);

                /*---- 2. Create DataTable ----*/
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("party_name", typeof(string)));
                dt.Columns.Add(new DataColumn("ship_to_address", typeof(string)));
                dt.Columns.Add(new DataColumn("ship_gstin_unique_id", typeof(string)));
                dt.Columns.Add(new DataColumn("ship_state", typeof(string)));
                dt.Columns.Add(new DataColumn("ship_state_code", typeof(string)));
                dt.Columns.Add(new DataColumn("bill_to_address", typeof(string)));
                dt.Columns.Add(new DataColumn("bill_gstin_unique_id", typeof(string)));
                dt.Columns.Add(new DataColumn("bill_state", typeof(string)));
                dt.Columns.Add(new DataColumn("bill_state_code", typeof(string)));
                dt.Columns.Add(new DataColumn("place_of_supply", typeof(string)));
                dt.Columns.Add(new DataColumn("invoice_no", typeof(string)));
                dt.Columns.Add(new DataColumn("credit_note_no", typeof(string)));
                dt.Columns.Add(new DataColumn("invoice_date", typeof(string)));
                dt.Columns.Add(new DataColumn("po_no", typeof(string)));
                dt.Columns.Add(new DataColumn("po_date", typeof(string)));
                dt.Columns.Add(new DataColumn("desc_resource_name", typeof(string)));
                dt.Columns.Add(new DataColumn("desc_billing_period", typeof(string)));
                dt.Columns.Add(new DataColumn("billable_days", typeof(string)));
                dt.Columns.Add(new DataColumn("quantity", typeof(string)));
                dt.Columns.Add(new DataColumn("quantity_type", typeof(string)));
                dt.Columns.Add(new DataColumn("rate", typeof(string)));
                dt.Columns.Add(new DataColumn("rate_type", typeof(string)));
                dt.Columns.Add(new DataColumn("amount", typeof(string)));
                dt.Columns.Add(new DataColumn("igst", typeof(string)));
                dt.Columns.Add(new DataColumn("sgst", typeof(string)));
                dt.Columns.Add(new DataColumn("cgst", typeof(string)));
                dt.Columns.Add(new DataColumn("tax_total", typeof(string)));
                dt.Columns.Add(new DataColumn("invoice_total", typeof(string)));
                dt.Columns.Add(new DataColumn("amount_in_words", typeof(string)));
                dt.Columns.Add(new DataColumn("invoice_file_name", typeof(string)));
                dt.Columns.Add(new DataColumn("doj", typeof(string)));
                dt.Columns.Add(new DataColumn("location", typeof(string)));
                dt.Columns.Add(new DataColumn("project_no", typeof(string)));
                dt.Columns.Add(new DataColumn("receipt_no", typeof(string)));
                dt.Columns.Add(new DataColumn("po_line_item", typeof(string)));
                dt.Columns.Add(new DataColumn("pbrs_client_id", typeof(string)));
                dt.Columns.Add(new DataColumn("update_on", typeof(string)));
                dt.Columns.Add(new DataColumn("sno", typeof(string)));
                dt.Columns.Add(new DataColumn("client_name", typeof(string)));
                dt.Columns.Add(new DataColumn("bill_start_date", typeof(string)));
                dt.Columns.Add(new DataColumn("bill_end_date", typeof(string)));
                dt.Columns.Add(new DataColumn("period", typeof(string)));
                dt.Columns.Add(new DataColumn("no_of_hours", typeof(string)));
                dt.Columns.Add(new DataColumn("due_date", typeof(string)));
                dt.Columns.Add(new DataColumn("currency", typeof(string)));
                dt.Columns.Add(new DataColumn("amount_received", typeof(string)));
                dt.Columns.Add(new DataColumn("outstanding", typeof(string)));
                dt.Columns.Add(new DataColumn("narration", typeof(string)));
                dt.Columns.Add(new DataColumn("employee_id", typeof(string)));
                dt.Columns.Add(new DataColumn("client_id", typeof(string)));
                dt.Columns.Add(new DataColumn("tax_zone", typeof(string)));
                dt.Columns.Add(new DataColumn("billed_by", typeof(string)));
                dt.Columns.Add(new DataColumn("remarks", typeof(string)));
                dt.Columns.Add(new DataColumn("shift_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("service_amount", typeof(string)));
                dt.Columns.Add(new DataColumn("system_invoice_no", typeof(string)));
                dt.Columns.Add(new DataColumn("invoice_type", typeof(string)));
                dt.Columns.Add(new DataColumn("sub_total", typeof(string)));
                dt.Columns.Add(new DataColumn("nature_of_business", typeof(string)));
                dt.Columns.Add(new DataColumn("sac_code", typeof(string)));
                dt.Columns.Add(new DataColumn("issue_date", typeof(string)));
                dt.Columns.Add(new DataColumn("irn_no", typeof(string)));
                dt.Columns.Add(new DataColumn("irn_date", typeof(string)));
                dt.Columns.Add(new DataColumn("qr_code", typeof(string)));
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
                            dr["party_name"] = worksheet.Cells[row, 1].Value == null ? "-" : worksheet.Cells[row, 1].Value.ToString().Trim();
                            dr["ship_to_address"] = worksheet.Cells[row, 2].Value == null ? "-" : worksheet.Cells[row, 2].Value.ToString().Trim();
                            dr["ship_gstin_unique_id"] = worksheet.Cells[row, 3].Value == null ? "-" : worksheet.Cells[row, 3].Value.ToString().Trim();
                            dr["ship_state"] = worksheet.Cells[row, 4].Value == null ? "-" : worksheet.Cells[row, 4].Value.ToString().Trim();
                            dr["ship_state_code"] = worksheet.Cells[row, 5].Value == null ? "-" : worksheet.Cells[row, 5].Value.ToString().Trim();
                            dr["bill_to_address"] = worksheet.Cells[row, 6].Value == null ? "-" : worksheet.Cells[row, 6].Value.ToString().Trim();
                            dr["bill_gstin_unique_id"] = worksheet.Cells[row, 7].Value == null ? "-" : worksheet.Cells[row, 7].Value.ToString().Trim();
                            dr["bill_state"] = worksheet.Cells[row, 8].Value == null ? "-" : worksheet.Cells[row, 8].Value.ToString().Trim();
                            dr["bill_state_code"] = worksheet.Cells[row, 9].Value == null ? "-" : worksheet.Cells[row, 9].Value.ToString().Trim();
                            dr["place_of_supply"] = worksheet.Cells[row, 10].Value == null ? "-" : worksheet.Cells[row, 10].Value.ToString().Trim();
                            dr["invoice_no"] = worksheet.Cells[row, 11].Value == null ? "-" : worksheet.Cells[row, 11].Value.ToString().Trim();
                            dr["invoice_date"] = worksheet.Cells[row, 12].Value == null ? "-" : worksheet.Cells[row, 12].Value.ToString().Trim();
                            dr["po_no"] = worksheet.Cells[row, 13].Value == null ? "-" : worksheet.Cells[row, 13].Value.ToString().Trim();
                            dr["po_date"] = worksheet.Cells[row, 14].Value == null ? "-" : worksheet.Cells[row, 14].Value.ToString().Trim();
                            dr["desc_resource_name"] = worksheet.Cells[row, 15].Value == null ? "-" : worksheet.Cells[row, 15].Value.ToString().Trim();
                            dr["desc_billing_period"] = worksheet.Cells[row, 16].Value == null ? "-" : worksheet.Cells[row, 16].Value.ToString().Trim();
                            dr["billable_days"] = worksheet.Cells[row, 17].Value == null ? "-" : worksheet.Cells[row, 17].Value.ToString().Trim();
                            dr["quantity"] = worksheet.Cells[row, 18].Value == null ? "-" : worksheet.Cells[row, 18].Value.ToString().Trim();
                            dr["quantity_type"] = worksheet.Cells[row, 19].Value == null ? "-" : worksheet.Cells[row, 19].Value.ToString().Trim();
                            dr["rate"] = worksheet.Cells[row, 20].Value == null ? "-" : worksheet.Cells[row, 20].Value.ToString().Trim();
                            dr["rate_type"] = worksheet.Cells[row, 21].Value == null ? "-" : worksheet.Cells[row, 21].Value.ToString().Trim();
                            dr["amount"] = worksheet.Cells[row, 22].Value == null ? "-" : worksheet.Cells[row, 22].Value.ToString().Trim();
                            dr["igst"] = worksheet.Cells[row, 23].Value == null ? "-" : worksheet.Cells[row, 23].Value.ToString().Trim();
                            dr["sgst"] = worksheet.Cells[row, 24].Value == null ? "-" : worksheet.Cells[row, 24].Value.ToString().Trim();
                            dr["cgst"] = worksheet.Cells[row, 25].Value == null ? "-" : worksheet.Cells[row, 25].Value.ToString().Trim();
                            dr["tax_total"] = worksheet.Cells[row, 26].Value == null ? "-" : worksheet.Cells[row, 26].Value.ToString().Trim();
                            dr["invoice_total"] = worksheet.Cells[row, 27].Value == null ? "-" : worksheet.Cells[row, 27].Value.ToString().Trim();
                            dr["amount_in_words"] = worksheet.Cells[row, 28].Value == null ? "-" : worksheet.Cells[row, 28].Value.ToString().Trim();
                            dr["invoice_file_name"] = worksheet.Cells[row, 29].Value == null ? "-" : worksheet.Cells[row, 29].Value.ToString().Trim();
                            dr["doj"] = worksheet.Cells[row, 30].Value == null ? "-" : worksheet.Cells[row, 30].Value.ToString().Trim();
                            dr["location"] = worksheet.Cells[row, 31].Value == null ? "-" : worksheet.Cells[row, 31].Value.ToString().Trim();
                            dr["project_no"] = worksheet.Cells[row, 32].Value == null ? "-" : worksheet.Cells[row, 32].Value.ToString().Trim();
                            dr["receipt_no"] = worksheet.Cells[row, 33].Value == null ? "-" : worksheet.Cells[row, 33].Value.ToString().Trim();
                            dr["po_line_item"] = worksheet.Cells[row, 34].Value == null ? "-" : worksheet.Cells[row, 34].Value.ToString().Trim();
                            dr["pbrs_client_id"] = worksheet.Cells[row, 35].Value == null ? "-" : worksheet.Cells[row, 35].Value.ToString().Trim();
                            dr["update_on"] = worksheet.Cells[row, 36].Value == null ? "-" : worksheet.Cells[row, 36].Value.ToString().Trim();
                            dr["sno"] = worksheet.Cells[row, 37].Value == null ? "-" : worksheet.Cells[row, 37].Value.ToString().Trim();
                            dr["client_name"] = worksheet.Cells[row, 38].Value == null ? "-" : worksheet.Cells[row, 38].Value.ToString().Trim();
                            dr["bill_start_date"] = worksheet.Cells[row, 39].Value == null ? "-" : worksheet.Cells[row, 39].Value.ToString().Trim();
                            dr["bill_end_date"] = worksheet.Cells[row, 40].Value == null ? "-" : worksheet.Cells[row, 40].Value.ToString().Trim();
                            dr["period"] = worksheet.Cells[row, 41].Value == null ? "-" : worksheet.Cells[row, 41].Value.ToString().Trim();
                            dr["no_of_hours"] = worksheet.Cells[row, 42].Value == null ? "-" : worksheet.Cells[row, 42].Value.ToString().Trim();
                            dr["due_date"] = worksheet.Cells[row, 43].Value == null ? "-" : worksheet.Cells[row, 43].Value.ToString().Trim();
                            dr["currency"] = worksheet.Cells[row, 44].Value == null ? "-" : worksheet.Cells[row, 44].Value.ToString().Trim();
                            dr["amount_received"] = worksheet.Cells[row, 45].Value == null ? "-" : worksheet.Cells[row, 45].Value.ToString().Trim();
                            dr["outstanding"] = worksheet.Cells[row, 46].Value == null ? "-" : worksheet.Cells[row, 46].Value.ToString().Trim();
                            dr["narration"] = worksheet.Cells[row, 47].Value == null ? "-" : worksheet.Cells[row, 47].Value.ToString().Trim();
                            dr["employee_id"] = worksheet.Cells[row, 48].Value == null ? "-" : worksheet.Cells[row, 48].Value.ToString().Trim();
                            dr["client_id"] = worksheet.Cells[row, 49].Value == null ? "-" : worksheet.Cells[row, 49].Value.ToString().Trim();
                            dr["tax_zone"] = worksheet.Cells[row, 50].Value == null ? "-" : worksheet.Cells[row, 50].Value.ToString().Trim();
                            dr["billed_by"] = worksheet.Cells[row, 51].Value == null ? "-" : worksheet.Cells[row, 51].Value.ToString().Trim();
                            dr["remarks"] = worksheet.Cells[row, 52].Value == null ? "-" : worksheet.Cells[row, 52].Value.ToString().Trim();
                            dr["shift_amount"] = worksheet.Cells[row, 53].Value == null ? "-" : worksheet.Cells[row, 53].Value.ToString().Trim();
                            dr["service_amount"] = worksheet.Cells[row, 54].Value == null ? "-" : worksheet.Cells[row, 54].Value.ToString().Trim();
                            dr["system_invoice_no"] = worksheet.Cells[row, 55].Value == null ? "-" : worksheet.Cells[row, 55].Value.ToString().Trim();
                            dr["nature_of_business"] = worksheet.Cells[row, 56].Value == null ? "-" : worksheet.Cells[row, 56].Value.ToString().Trim();
                            dr["sac_code"] = worksheet.Cells[row, 57].Value == null ? "-" : worksheet.Cells[row, 57].Value.ToString().Trim();
                            dr["irn_no"] = worksheet.Cells[row, 58].Value == null ? "-" : worksheet.Cells[row, 58].Value.ToString().Trim();
                            dr["irn_date"] = worksheet.Cells[row, 59].Value == null ? "-" : worksheet.Cells[row, 59].Value.ToString().Trim();
                            dr["qr_code"] = worksheet.Cells[row, 60].Value == null ? "-" : worksheet.Cells[row, 60].Value.ToString().Trim();
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
                    objbulk.DestinationTableName = "billing_spoc";

                    objbulk.ColumnMappings.Add("party_name", "party_name");
                    objbulk.ColumnMappings.Add("ship_to_address", "ship_to_address");
                    objbulk.ColumnMappings.Add("ship_gstin_unique_id", "ship_gstin_unique_id");
                    objbulk.ColumnMappings.Add("ship_state", "ship_state");
                    objbulk.ColumnMappings.Add("ship_state_code", "ship_state_code");
                    objbulk.ColumnMappings.Add("bill_to_address", "bill_to_address");
                    objbulk.ColumnMappings.Add("bill_gstin_unique_id", "bill_gstin_unique_id");
                    objbulk.ColumnMappings.Add("bill_state", "bill_state");
                    objbulk.ColumnMappings.Add("bill_state_code", "bill_state_code");
                    objbulk.ColumnMappings.Add("place_of_supply", "place_of_supply");
                    objbulk.ColumnMappings.Add("invoice_no", "invoice_no");
                    objbulk.ColumnMappings.Add("credit_note_no", "credit_note_no");
                    objbulk.ColumnMappings.Add("invoice_date", "invoice_date");
                    objbulk.ColumnMappings.Add("po_no", "po_no");
                    objbulk.ColumnMappings.Add("po_date", "po_date");
                    objbulk.ColumnMappings.Add("desc_resource_name", "desc_resource_name");
                    objbulk.ColumnMappings.Add("desc_billing_period", "desc_billing_period");
                    objbulk.ColumnMappings.Add("billable_days", "billable_days");
                    objbulk.ColumnMappings.Add("quantity", "quantity");
                    objbulk.ColumnMappings.Add("quantity_type", "quantity_type");
                    objbulk.ColumnMappings.Add("rate", "rate");
                    objbulk.ColumnMappings.Add("rate_type", "rate_type");
                    objbulk.ColumnMappings.Add("amount", "amount");
                    objbulk.ColumnMappings.Add("igst", "igst");
                    objbulk.ColumnMappings.Add("sgst", "sgst");
                    objbulk.ColumnMappings.Add("cgst", "cgst");
                    objbulk.ColumnMappings.Add("tax_total", "tax_total");
                    objbulk.ColumnMappings.Add("invoice_total", "invoice_total");
                    objbulk.ColumnMappings.Add("amount_in_words", "amount_in_words");
                    objbulk.ColumnMappings.Add("invoice_file_name", "invoice_file_name");
                    objbulk.ColumnMappings.Add("doj", "doj");
                    objbulk.ColumnMappings.Add("location", "location");
                    objbulk.ColumnMappings.Add("project_no", "project_no");
                    objbulk.ColumnMappings.Add("receipt_no", "receipt_no");
                    objbulk.ColumnMappings.Add("po_line_item", "po_line_item");
                    objbulk.ColumnMappings.Add("pbrs_client_id", "pbrs_client_id");
                    objbulk.ColumnMappings.Add("update_on", "update_on");
                    objbulk.ColumnMappings.Add("sno", "sno");
                    objbulk.ColumnMappings.Add("client_name", "client_name");
                    objbulk.ColumnMappings.Add("bill_start_date", "bill_start_date");
                    objbulk.ColumnMappings.Add("bill_end_date", "bill_end_date");
                    objbulk.ColumnMappings.Add("period", "period");
                    objbulk.ColumnMappings.Add("no_of_hours", "no_of_hours");
                    objbulk.ColumnMappings.Add("due_date", "due_date");
                    objbulk.ColumnMappings.Add("currency", "currency");
                    objbulk.ColumnMappings.Add("amount_received", "amount_received");
                    objbulk.ColumnMappings.Add("outstanding", "outstanding");
                    objbulk.ColumnMappings.Add("narration", "narration");
                    objbulk.ColumnMappings.Add("employee_id", "employee_id");
                    objbulk.ColumnMappings.Add("client_id", "client_id");
                    objbulk.ColumnMappings.Add("tax_zone", "tax_zone");
                    objbulk.ColumnMappings.Add("billed_by", "billed_by");
                    objbulk.ColumnMappings.Add("remarks", "remarks");
                    objbulk.ColumnMappings.Add("shift_amount", "shift_amount");
                    objbulk.ColumnMappings.Add("service_amount", "service_amount");
                    objbulk.ColumnMappings.Add("system_invoice_no", "system_invoice_no");
                    objbulk.ColumnMappings.Add("invoice_type", "invoice_type");
                    objbulk.ColumnMappings.Add("sub_total", "sub_total");
                    objbulk.ColumnMappings.Add("nature_of_business", "nature_of_business");
                    objbulk.ColumnMappings.Add("sac_code", "sac_code");
                    objbulk.ColumnMappings.Add("issue_date", "issue_date");
                    objbulk.ColumnMappings.Add("irn_no", "irn_no");
                    objbulk.ColumnMappings.Add("irn_date", "irn_date");
                    objbulk.ColumnMappings.Add("qr_code", "qr_code");
                    objbulk.ColumnMappings.Add("file_uploader_id", "file_uploader_id");
                    objbulk.ColumnMappings.Add("timestamp", "timestamp");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(dt);
                    con.Close();
                }

                /*--- 6.Return With Success Status ---*/
                status = true;
                _logger.LogInformation("uploadBillingFile Method Success");
                return status;

            }
            catch (Exception ex)
            {
                status = false;
                _logger.LogInformation("uploadBillingFile -------- " + ex.Message + ex.StackTrace);
                return status;
            }
        }

        public bool sendBillingMail()
        {
            bool returnStatus = false;
            string[] toMail = { };
            string[] ccMail = { };
            var clientName = "";
            var htmlBody = "";
            var signature = "";
            var resource = "";
            var mailSubject = "";
            var invoice_total = "";
            DataTable dt_billingSPOCLog = new DataTable();
            DataTable dt_billingSPOCLog_duplicate = new DataTable();
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string empID = _httpContextAccessor.HttpContext.Session.GetString("emp_id");
            var billingInvoicePath = _configuration.GetValue<string>("BillingInvoicePath");
            string webRootPath = _HostEnvironment.WebRootPath;

            if (username == null || empID == null)
            {
                returnStatus = false;
                return returnStatus;
            }

            /*----- 0. Log Info Into File -----*/
            _logger.LogInformation("Billing SPOC Mail Sent Started");

            try
            {
                /*---------- Get All Data Of Billing SPOC ----------*/
                dt_billingSPOCLog = getBillingSPOCLog(username);
                if (dt_billingSPOCLog.Rows.Count > 0)
                {
                    /*---- Get Distinct Invoice No List From Datatable -----*/
                    var distinctInvoiceNo = (from spoc in dt_billingSPOCLog.AsEnumerable()
                                             select new
                                             {
                                                 invoiceNo = spoc.Field<string>("invoice_no")
                                             }).Distinct().ToList();

                    /*--- Loop Through Each Invoice In List ---*/
                    foreach (var invoice_no in distinctInvoiceNo)
                    {
                        /*--- Clear Data---*/
                        toMail = new string[] { };
                        ccMail = new string[] { };


                        /*---- Get DataRows Of Particular Invoice No ----*/
                        dt_billingSPOCLog_duplicate = dt_billingSPOCLog.AsEnumerable().Where(r => r.Field<string>("invoice_no") == invoice_no.invoiceNo.ToString()).CopyToDataTable();

                        /*---- Create HTML For Mail Body ------*/
                        if (dt_billingSPOCLog_duplicate.Rows.Count > 0)
                        {
                            /*--- Check If Attachment Exists---*/
                            /*--- Only Then Proceed To Mail Otherwise Skip Iteration ---*/
                            string pbrsClientID = dt_billingSPOCLog_duplicate.Rows[0]["pbrs_client_id"].ToString();
                            string invoiceNoWithoutSC = Regex.Replace(invoice_no.invoiceNo, @"(\s+|@|&|'|\(|\)|\\|\/|<|>|#)", "");

                            string attachmentPath = $@"{billingInvoicePath}\\{empID}\\{pbrsClientID}\\{invoiceNoWithoutSC}.pdf";
                            //string attachmentPath = Path.Combine($@"wwwroot/{billingInvoicePath}/{empID}/{pbrsClientID}/", "" + invoiceNoWithoutSC + ".pdf");


                            if (File.Exists(attachmentPath))
                            {
                                clientName = dt_billingSPOCLog_duplicate.Rows[0]["party_name"].ToString();

                                /*--- Check If Client Resource Is Single Or Multiple---*/
                                if (dt_billingSPOCLog_duplicate.Rows.Count == 1)
                                {
                                    resource = dt_billingSPOCLog_duplicate.Rows[0]["desc_resource_name"].ToString();
                                    invoice_total = dt_billingSPOCLog_duplicate.Rows[0]["invoice_total"].ToString();

                                    /*--- Create Mail Subject---*/
                                    mailSubject = $"{clientName} - {invoice_no.invoiceNo} - {resource}";
                                }
                                else
                                {
                                    resource = dt_billingSPOCLog_duplicate.Rows.Count.ToString();
                                    invoice_total = dt_billingSPOCLog_duplicate.AsEnumerable().Sum(x => x.Field<double>("invoice_total")).ToString();

                                    /*--- Create Mail Subject---*/
                                    mailSubject = $"{clientName} - {invoice_no.invoiceNo} - {resource} Resources";
                                }

                                /*--- Create Html Body---*/
                                htmlBody = @"<p style=""font-family:Calibri;font-size:15;"">Hi Team/Sir/Madam,</p></n><p style=""font-family:Calibri;font-size:15;"">Please find the attached invoice raised against below details. Request you to acknowledge the same and process for payment.</p></br>";
                                htmlBody += @"<table style=""font-family:Calibri;font-size:12;border-collapse: collapse;width: 100%;"">";
                                htmlBody += @"<tr>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Invoice No</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Invoice Date</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">PO No</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">PO Date</th>
		                                <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Resource Name</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Billing Period</th>
                                        <th style=""border: 1px solid #dddddd;text-align: center;padding: 8px;background-color: #dddddd;"">Invoice Total</th>
	                              </tr>";
                                htmlBody += $@"<tr>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dt_billingSPOCLog_duplicate.Rows[0]["invoice_no"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dt_billingSPOCLog_duplicate.Rows[0]["invoice_date"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dt_billingSPOCLog_duplicate.Rows[0]["po_no"].ToString()}</td>
                                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{dt_billingSPOCLog_duplicate.Rows[0]["po_date"].ToString()}</td>
		                                <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{resource}</td>
                                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">
                                    {dt_billingSPOCLog_duplicate.Rows[0]["desc_billing_period"].ToString()}</td>
                                        <td style=""border: 1px solid #dddddd;text-align: center;padding: 8px;"">{String.Format(new CultureInfo("en-IN", false), "{0:n}", Convert.ToDouble(invoice_total))}</td>
	                                </tr>";
                                htmlBody += @"</table>";
                                htmlBody += @"<br><br>";

                                /*---- Add Signature In Mail -----*/
                                string signaturePath = Path.Combine("wwwroot/Templates/Billing/", "" + empID + ".html");
                                if (File.Exists(signaturePath))
                                {
                                    using (StreamReader streamReader = new StreamReader(signaturePath))
                                    {
                                        signature = streamReader.ReadToEnd();
                                    }
                                    htmlBody += signature;
                                }

                                /*---- Set To and CC ----*/
                                toMail = dt_billingSPOCLog_duplicate.Rows[0]["to_mail"].ToString().Split(';');
                                ccMail = dt_billingSPOCLog_duplicate.Rows[0]["cc_mail"].ToString().Split(';');


                                /*--- Send Mail & Delete Invoice----*/
                                try
                                {
                                    sendMail(username, toMail, ccMail, htmlBody, mailSubject, attachmentPath);
                                    _logger.LogInformation("Billing Mail Finished For ---- " + invoice_no.invoiceNo);
                                    System.IO.File.Delete(attachmentPath);
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogInformation("Billing Mail Send Error For ---- " +invoice_no.invoiceNo + ex.Message + ex.StackTrace);
                                }



                                /*---- Update Invoice No Send Status ----*/
                                /*--- So that Doesn't mail Duplicate Again---*/
                                int sentUpdateStatus = updateInvoiceSendStatus(invoice_no.invoiceNo);
                                _logger.LogInformation("Updated Sent Status Of ---- " + invoice_no.invoiceNo);
                            }
                            else
                            {
                                /*----- 0. Log Info Into File -----*/
                                _logger.LogInformation("Attachment Not Found For Invoice No " + invoice_no.invoiceNo + ". Skipped. The path is "+ attachmentPath);
                            }
                        }
                    }

                    /*---- Set Return Status---*/
                    returnStatus = true;

                    /*----- 0. Log Info Into File -----*/
                    _logger.LogInformation("Billing SPOC Mail Sent Complete For All");
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
                _logger.LogInformation("sendBillingMail Method ---- " + ex.Message + ex.StackTrace);
                returnStatus = false;
            }
            return returnStatus;
        }

        public int deleteBillingSPOCTable(string loginID)
        {
            int status = 0;

            string sql_delete_billing_spoc_tbl = $"delete from billing_spoc where file_uploader_id = '{loginID}'";

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

        public dynamic getBillingSPOCLog(string loginId)
        {
            string sp_billing_log = "get_billing_spoc_log";
            DataTable dt_billing_log = new DataTable();

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_billing_log;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@login_id", loginId);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlDataAdapter oda = new SqlDataAdapter(cmd);
                oda.Fill(dt_billing_log);
            }

            return dt_billing_log;
        }

        public int updateInvoiceSendStatus(string invoiceNo)
        {
            int status = 0;

            string sql_update_invoice_send_status = $"update billing_spoc set sent_status='true' where invoice_no='{invoiceNo}'";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_update_invoice_send_status;
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

        public void sendMail(string Sender, string[] mailTo, string[] CCTo, string htmlBody, string subject, string attachmentPath)
        {
            var host = _configuration.GetValue<string>("MailSettings:Host");
            var port = _configuration.GetValue<int>("MailSettings:Port");
            var username = _configuration.GetValue<string>("MailSettings:UserName");
            var password = _configuration.GetValue<string>("MailSettings:Password");
            //var Sender = _configuration.GetValue<string>("MailSettings:Sender");

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

        public bool uploadMaillingAddresses(IFormFile file)
        {
            bool status = false;
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

                /*---- 1. Truncate 'mailling_address' Table---*/
                int truncateStatus = truncateMaillingAddressTable();

                /*---- 2. Create DataTable ----*/
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("resource_name", typeof(string)));
                dt.Columns.Add(new DataColumn("employee_id", typeof(string)));
                dt.Columns.Add(new DataColumn("client_name", typeof(string)));
                dt.Columns.Add(new DataColumn("pbrs_id", typeof(string)));
                dt.Columns.Add(new DataColumn("to_mail", typeof(string)));
                dt.Columns.Add(new DataColumn("cc_mail", typeof(string)));

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
                            dr["resource_name"] = worksheet.Cells[row, 1].Value == null ? "-" : worksheet.Cells[row, 1].Value.ToString().Trim();
                            dr["employee_id"] = worksheet.Cells[row, 2].Value == null ? "-" : worksheet.Cells[row, 2].Value.ToString().Trim();
                            dr["client_name"] = worksheet.Cells[row, 3].Value == null ? "-" : worksheet.Cells[row, 3].Value.ToString().Trim();
                            dr["pbrs_id"] = worksheet.Cells[row, 4].Value == null ? "-" : worksheet.Cells[row, 4].Value.ToString().Trim();
                            dr["to_mail"] = worksheet.Cells[row, 5].Value == null ? "-" : worksheet.Cells[row, 5].Value.ToString().Trim();
                            dr["cc_mail"] = worksheet.Cells[row, 6].Value == null ? "-" : worksheet.Cells[row, 6].Value.ToString().Trim();

                            dt.Rows.Add(dr);
                        }
                    }
                }

                /*--- 5. Bulk Insert In 'mailling_address' Table ---*/
                using (SqlConnection con = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    objbulk.DestinationTableName = "mailling_address";

                    objbulk.ColumnMappings.Add("resource_name", "resource_name");
                    objbulk.ColumnMappings.Add("employee_id", "employee_id");
                    objbulk.ColumnMappings.Add("client_name", "client_name");
                    objbulk.ColumnMappings.Add("pbrs_id", "pbrs_id");
                    objbulk.ColumnMappings.Add("to_mail", "to_mail");
                    objbulk.ColumnMappings.Add("cc_mail", "cc_mail");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(dt);
                    con.Close();
                }

                /*--- 6.Return With Success Status ---*/
                status = true;
                _logger.LogInformation("Mail Address File Upload Success");
                return status;
            }
            catch (Exception ex)
            {
                status = false;
                _logger.LogInformation("uploadMaillingAddresses -------- " + ex.Message + ex.StackTrace);
                return status;
            }
        }

        public int truncateMaillingAddressTable()
        {
            int status = 0;

            string sql_truncate_mail_addrs_tbl = "truncate table mailling_address";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_truncate_mail_addrs_tbl;
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
    }
}
