using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Extensions.Configuration;
using System;
using EITFinance.Models;
using EITFinance.Services;
using EITFinance.Repositories;
using Microsoft.Extensions.Logging;

namespace EITFinance.Controllers
{
    public class CollectionSummaryController : Controller
    {
        private IConfiguration _configuration;
        ICollectionSummaryService _collectionSummary;
        private readonly ILogger<SchedulerRepository> _logger;
        IHttpContextAccessor _httpContextAccessor;
        ILoginService _loginService;
        public CollectionSummaryController(IConfiguration Configuration, ICollectionSummaryService CollectionSummary 
                                           , IHttpContextAccessor httpContextAccessor, ILoginService loginService
                                           , ILogger<SchedulerRepository> logger)
        {
            _configuration = Configuration;
            _collectionSummary = CollectionSummary;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
            _loginService = loginService;
        }

        public IActionResult Index()
        {
            //string cookie = Request.Cookies["Role"];
            //if (cookie == null)
            //{
            //    return RedirectToAction("Index", "Login", new { message = "You-are-not-Authorised." });
            //}
            //else
            //{
            //    ViewBag.successStatus = TempData["successStatus"];
            //    return View();
            //}
            string username = _httpContextAccessor.HttpContext.Session.GetString("username");
            string ModuleCollectionSummary = _configuration.GetValue<string>("ModuleCollectionSummary");
            if (username == null)
            {
                return RedirectToAction("Index", "Login");
            }
            else
            {
                DataTable dt_authorized_status = _loginService.getAuthorizationStatus(username, ModuleCollectionSummary);
                if (dt_authorized_status.Rows.Count == 0)
                {
                    return RedirectToAction("Index", "Dashboard");
                }
            }


            ViewBag.successStatus = TempData["successStatus"];
            return View();
        }

        [RequestSizeLimit(10000000)]
        public IActionResult UploadCollectionSummary(IFormFile file)
        {
            string cookie = Request.Cookies["Role"];
            if (cookie == null)
            {
                return RedirectToAction("Index", "Login", new { message = "You are not Authorised. Please Login." });
            }
            var datetime = DateTime.Now.ToString("ddMMyyyyhhmm");
            try
            {
                /*---- 0. Check File Extension ----*/
                if (file != null)
                {
                    var extension = Path.GetExtension(file.FileName);
                    if (extension.ToLower() != ".xlsx")
                    {
                        TempData["successStatus"] = false;
                        return RedirectToAction("Index");
                    }
                }
                else
                {
                    TempData["successStatus"] = false;
                    return RedirectToAction("Index");
                }

                /*---- 1. Truncate 'collection_summary_staging' Table---*/
                int truncateStatus = _collectionSummary.truncateCollectionSummaryStagingTable();


                /*---- 2. Create DataTable ----*/
                DataTable dt = new DataTable();
                dt.Columns.Add(new DataColumn("client_name", typeof(string)));
                dt.Columns.Add(new DataColumn("cheque_date", typeof(string)));
                dt.Columns.Add(new DataColumn("cheque_details", typeof(string)));
                dt.Columns.Add(new DataColumn("payment_received_date", typeof(DateTime)));
                dt.Columns.Add(new DataColumn("amount", typeof(string)));
                dt.Columns.Add(new DataColumn("currency", typeof(string)));
                dt.Columns.Add(new DataColumn("currency_rate", typeof(string)));
                dt.Columns.Add(new DataColumn("amount_in_inr", typeof(string)));
                dt.Columns.Add(new DataColumn("payment_type", typeof(string)));
                dt.Columns.Add(new DataColumn("payment_details", typeof(string)));
                dt.Columns.Add(new DataColumn("remittance_status", typeof(string)));
                dt.Columns.Add(new DataColumn("email_date_for_remittance", typeof(string)));
                dt.Columns.Add(new DataColumn("accounts_updation_date", typeof(string)));
                dt.Columns.Add(new DataColumn("remarks", typeof(string)));
                dt.Columns.Add(new DataColumn("ad", typeof(string)));
                dt.Columns.Add(new DataColumn("bd", typeof(string)));
                dt.Columns.Add(new DataColumn("client_spoc", typeof(string)));
                dt.Columns.Add(new DataColumn("billing_spoc", typeof(string)));
                dt.Columns.Add(new DataColumn("type", typeof(string)));
                dt.Columns.Add(new DataColumn("coll_month_year", typeof(string)));
                dt.Columns.Add(new DataColumn("week", typeof(string)));
                dt.Columns.Add(new DataColumn("year", typeof(string)));
                dt.Columns.Add(new DataColumn("coll_month", typeof(string)));
                dt.Columns.Add(new DataColumn("adj_month", typeof(string)));
                dt.Columns.Add(new DataColumn("to_send", typeof(string)));
                dt.Columns.Add(new DataColumn("cc", typeof(string)));
                dt.Columns.Add(new DataColumn("mail_remarks", typeof(string)));
                dt.Columns.Add(new DataColumn("upload_date", typeof(string)));

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
                                dr["client_name"] = worksheet.Cells[row, 2].Value == null ? "-" : worksheet.Cells[row, 2].Value.ToString().Trim();
                                dr["cheque_date"] = worksheet.Cells[row, 3].Value == null ? "-" : worksheet.Cells[row, 3].Value.ToString().Trim();
                                dr["cheque_details"] = worksheet.Cells[row, 4].Value == null ? "-" : worksheet.Cells[row, 4].Value.ToString().Trim();
                                dr["payment_received_date"] = worksheet.Cells[row, 5].Value == null ? "-" : worksheet.Cells[row, 5].Value.ToString().Trim();
                                dr["amount"] = worksheet.Cells[row, 6].Value == null ? "-" : worksheet.Cells[row, 6].Value.ToString().Trim();
                                dr["currency"] = worksheet.Cells[row, 7].Value == null ? "-" : worksheet.Cells[row, 7].Value.ToString().Trim();
                                dr["currency_rate"] = worksheet.Cells[row, 8].Value == null ? "-" : worksheet.Cells[row, 8].Value.ToString().Trim();
                                dr["amount_in_inr"] = worksheet.Cells[row, 9].Value == null ? "-" : worksheet.Cells[row, 9].Value.ToString().Trim();
                                dr["payment_type"] = worksheet.Cells[row, 10].Value == null ? "-" : worksheet.Cells[row, 10].Value.ToString().Trim();
                                dr["payment_details"] = worksheet.Cells[row, 11].Value == null ? "-" : worksheet.Cells[row, 11].Value.ToString().Trim();
                                dr["remittance_status"] = worksheet.Cells[row, 12].Value == null ? "-" : worksheet.Cells[row, 12].Value.ToString().Trim();
                                dr["email_date_for_remittance"] = worksheet.Cells[row, 13].Value == null ? "-" : worksheet.Cells[row, 13].Value.ToString().Trim();
                                dr["accounts_updation_date"] = worksheet.Cells[row, 14].Value == null ? "-" : worksheet.Cells[row, 14].Value.ToString().Trim();
                                dr["remarks"] = worksheet.Cells[row, 15].Value == null ? "-" : worksheet.Cells[row, 15].Value.ToString().Trim();
                                dr["ad"] = worksheet.Cells[row, 16].Value == null ? "-" : worksheet.Cells[row, 16].Value.ToString().Trim();
                                dr["bd"] = worksheet.Cells[row, 17].Value == null ? "-" : worksheet.Cells[row, 17].Value.ToString().Trim();
                                dr["client_spoc"] = worksheet.Cells[row, 18].Value == null ? "-" : worksheet.Cells[row, 18].Value.ToString().Trim();
                                dr["billing_spoc"] = worksheet.Cells[row, 19].Value == null ? "-" : worksheet.Cells[row, 19].Value.ToString().Trim();
                                dr["type"] = worksheet.Cells[row, 20].Value == null ? "-" : worksheet.Cells[row, 20].Value.ToString().Trim();
                                dr["coll_month_year"] = worksheet.Cells[row, 21].Value == null ? "-" : worksheet.Cells[row, 21].Value.ToString().Trim();
                                dr["week"] = worksheet.Cells[row, 22].Value == null ? "-" : worksheet.Cells[row, 22].Value.ToString().Trim();
                                dr["year"] = worksheet.Cells[row, 23].Value == null ? "-" : worksheet.Cells[row, 23].Value.ToString().Trim();
                                dr["coll_month"] = worksheet.Cells[row, 24].Value == null ? "-" : worksheet.Cells[row, 24].Value.ToString().Trim();
                                dr["adj_month"] = worksheet.Cells[row, 25].Value == null ? "-" : worksheet.Cells[row, 25].Value.ToString().Trim();
                                dr["to_send"] = worksheet.Cells[row, 26].Value == null ? "-" : worksheet.Cells[row, 26].Value.ToString().Trim();
                                dr["cc"] = worksheet.Cells[row, 27].Value == null ? "-" : worksheet.Cells[row, 27].Value.ToString().Trim();
                                dr["mail_remarks"] = worksheet.Cells[row, 28].Value == null ? "-" : worksheet.Cells[row, 28].Value.ToString().Trim();
                                dr["upload_date"] = datetime;

                                dt.Rows.Add(dr);
                            
                        }
                    }
                }


                /*--- 5. Bulk Insert In 'collection_summary_staging' Table ---*/
                using (SqlConnection con = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    objbulk.DestinationTableName = "collection_summary_staging";

                    objbulk.ColumnMappings.Add("client_name", "client_name");
                    objbulk.ColumnMappings.Add("cheque_date", "cheque_date");
                    objbulk.ColumnMappings.Add("cheque_details", "cheque_details");
                    objbulk.ColumnMappings.Add("payment_received_date", "payment_received_date");
                    objbulk.ColumnMappings.Add("amount", "amount");
                    objbulk.ColumnMappings.Add("currency", "currency");
                    objbulk.ColumnMappings.Add("currency_rate", "currency_rate");
                    objbulk.ColumnMappings.Add("amount_in_inr", "amount_in_inr");
                    objbulk.ColumnMappings.Add("payment_type", "payment_type");
                    objbulk.ColumnMappings.Add("payment_details", "payment_details");
                    objbulk.ColumnMappings.Add("remittance_status", "remittance_status");
                    objbulk.ColumnMappings.Add("email_date_for_remittance", "email_date_for_remittance");
                    objbulk.ColumnMappings.Add("accounts_updation_date", "accounts_updation_date");
                    objbulk.ColumnMappings.Add("remarks", "remarks");
                    objbulk.ColumnMappings.Add("ad", "ad");
                    objbulk.ColumnMappings.Add("bd", "bd");
                    objbulk.ColumnMappings.Add("client_spoc", "client_spoc");
                    objbulk.ColumnMappings.Add("billing_spoc", "billing_spoc");
                    objbulk.ColumnMappings.Add("type", "type");
                    objbulk.ColumnMappings.Add("coll_month_year", "coll_month_year");
                    objbulk.ColumnMappings.Add("week", "week");
                    objbulk.ColumnMappings.Add("year", "year");
                    objbulk.ColumnMappings.Add("coll_month", "coll_month");
                    objbulk.ColumnMappings.Add("adj_month", "adj_month");
                    objbulk.ColumnMappings.Add("to_send", "to_send");
                    objbulk.ColumnMappings.Add("cc", "cc");
                    objbulk.ColumnMappings.Add("mail_remarks", "mail_remarks");
                    objbulk.ColumnMappings.Add("upload_date", "upload_date");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(dt);
                    con.Close();
                }

                /*---- 6. Truncate 'collection_summary' Table---*/
                int truncateStatusColSum = _collectionSummary.truncateCollectionSummaryTable();

                /*---- 7. Merge into 'collection_summary' Main Table ---*/
                int mergeStatus = _collectionSummary.mergeFromStagingToMainTable();


                /*-- 8. Send Status To View To Show Alert --*/
                TempData["successStatus"] = true;
            }
            catch (Exception ex)
            {
                /*-- Send Status To View To Show Alert --*/
                TempData["successStatus"] = false;

                /*----- Log Exception Info Into File -----*/
                _logger.LogInformation(ex.StackTrace);
            }
            return RedirectToAction("Index");
        }
    }
}
