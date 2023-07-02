using EITFinance.Models;
using EITFinance.Models.Timesheet;
using EITFinance.Models.Timesheet.DTOs;
using EITFinance.Services;
using EITFinance.Utilities;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;

namespace EITFinance.Repositories
{
    public class TimeSheetRepository : ITimesheetService
    {
        IConfiguration _configuration;
        static SqlConnection conn = null;
        private readonly ILogger<SchedulerRepository> _logger;
        private IEmailSender _emailSender;
        IHttpContextAccessor _httpContextAccessor;

        public TimeSheetRepository(IConfiguration configuration, ILogger<SchedulerRepository> logger, IEmailSender emailSender, IHttpContextAccessor httpContextAccessor)
        {
            _configuration = configuration;
            _logger = logger;
            _emailSender = emailSender;
            _httpContextAccessor = httpContextAccessor;
        }
        public void TimesheetProcessor()
        {
            var schedulerFlag = _configuration.GetValue<bool>("Application:Scheduler");
            var startTime = _configuration.GetValue<string>("Application:SchedulerStartTimespan").Split(":");
            var endTime = _configuration.GetValue<string>("Application:SchedulerEndTimespan").Split(":");
            try
            {
                TimeSpan startAutoMail = new TimeSpan(int.Parse(startTime[0]), int.Parse(startTime[1]), int.Parse(startTime[2]));
                TimeSpan endAutoMail = new TimeSpan(int.Parse(endTime[0]), int.Parse(endTime[1]), int.Parse(endTime[2]));
                TimeSpan now = DateTime.Now.TimeOfDay;

                if ((now >= startAutoMail) && (now <= endAutoMail))
                {
                    _logger.LogInformation("Timesheet process has been started.");
                    _logger.LogInformation("geteting pending timesheets from sql table.");
                    DataTable dtTimesheets = GetTimesheets(1);  

                    if (dtTimesheets.Rows.Count > 0)
                    {
                        var dtClientName = dtTimesheets.DefaultView.ToTable(true, "clientname", "mailto", "mailcc", "pbrs_id");

                        foreach (DataRow dr in dtClientName.Rows)
                        {
                            try
                            {
                                string[] toRecepients = string.IsNullOrEmpty(dr["mailto"].ToString())
                                    ? _configuration.GetValue<string>("Application:defaultEmail").Split(';')
                                    : dr["mailto"].ToString().Split(';');

                                string[] CcRecepients = dr["mailcc"].ToString().Split(';');

                                //string[] toRecepients = { "shubhamay.kundu@in.experis.com" };
                                //string[] CcRecepients = { "sunil.kumar2@manpowergroup.com" } ;

                                IList<string> lstTimeSheets = dtTimesheets.AsEnumerable().Where(myRow => myRow.Field<string>("clientname") == dr["clientname"].ToString()).Select(item => string.Format("{0}", item["targetpath"])).ToList();
                                string timesheet = string.Join("</br>", lstTimeSheets);

                                string body = CreateBodyTemplate(timesheet);
                                body += FileHelper.GetSignature("EIT_Signature.html");

                                var clientName = dr["clientname"].ToString();

                                var message = new Message(toRecepients, CcRecepients, $"Timesheet Arrived – {clientName}", body);
                                _logger.LogInformation($"sending {clientName} mail to {toRecepients}");
                                _emailSender.SendEmail(message);
                                _logger.LogInformation($"updating timesheet status in database.");
                                updateStatus(dr["pbrs_id"].ToString());
                            }

                            catch (Exception ex)
                            {
                                _logger.LogError(ex.StackTrace);
                            }
                        }
                    }

                    _logger.LogInformation("Timesheet process has been completed.");
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation("TimesheetProcessor Method ---" + ex.StackTrace + "-----" + ex.Message);
            }

        }
        private string CreateBodyTemplate(string content)
        {

            string body = $@"<p style=""font-family:Calibri;font-size:12;color:black"">Hi Team,</p></n><p style=""font-family:Calibri;font-size:12;color:black"">The below are the Timesheets which were received today.</p></n><p>";
            body = $@"<p style=""font-family:Calibri;font-size:12;color:black"">Hi Team,</p></n><p style=""font-family:Calibri;font-size:12;color:black"">The below are the Timesheets which were received today.</br></br>";
            body += content;
            body += $@"</p></n><p style=""font-family:Calibri;font-size:12;color:black"">Thanks and Regards.</p>";
            return body;
        }
        public List<InsertTimesheetDTO> DirectoryScanner(string folderPath, string archiveFolderPath)
        {
            List<InsertTimesheetDTO> insertTimesheetDTO = new List<InsertTimesheetDTO>();
            string[] files = FileHelper.GetFiles(folderPath);

            if (files.Length == 0)
                return insertTimesheetDTO;

            foreach (string file in files)
            {
                string fileName = Path.GetFileName(file);
                string path = Path.GetDirectoryName(file);
                string relativePath = path.Replace(folderPath, "");
                string destinationPath = archiveFolderPath + "" + relativePath + "/" + fileName;

                if (!Directory.Exists(archiveFolderPath + "/" + relativePath))
                {
                    _logger.LogInformation($"Creating folder in archive path {archiveFolderPath + "/" + relativePath}");
                    Directory.CreateDirectory(archiveFolderPath + "/" + relativePath);
                }

                try
                {
                    _logger.LogInformation($"Coping file {file} to {destinationPath}");
                    File.Copy(file, destinationPath, true);
                    _logger.LogInformation($"Deleteing file {file}");
                    //File.Delete(file);
                    insertTimesheetDTO.Add(new InsertTimesheetDTO { clientName = relativePath.Split('\\')[1], sourcePath = file, targetPath = destinationPath });
                }
                catch (Exception ex)
                {
                    _logger.LogInformation(ex.Message);
                }
            }

            return insertTimesheetDTO;

        }
        public bool UploadMaillingAddresses(IFormFile file)
        {
            throw new System.NotImplementedException();
        }
        public DataTable GetTimesheets(int status)
        {
            DataTable dt = new DataTable();

            try
            {
                string sp_name = "getTimesheets";
                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = sp_name;
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@status", status);
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    SqlDataAdapter oda = new SqlDataAdapter(cmd);
                    oda.Fill(dt);
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation(ex.Message);
            }

            return dt;
        }

        public void Insert(List<InsertTimesheetDTO> timesheet)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlBulkCopy objbulk = new SqlBulkCopy(con);
                    objbulk.DestinationTableName = "TimeSheet";

                    objbulk.ColumnMappings.Add("clientName", "ClientName");
                    objbulk.ColumnMappings.Add("sourcePath", "SourcePath");
                    objbulk.ColumnMappings.Add("targetPath", "TargetPath");

                    if (con.State == ConnectionState.Closed) con.Open();
                    objbulk.WriteToServer(timesheet.AsDataTable());
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation(ex.Message);
            }
        }

        public int updateStatus(string pbrs_id)
        {
            int status = 0;
            try
            {
                string sql_update_status = $"update timesheet set status=2, modify_date=getdate()  where status=1 and pbrs_id='{pbrs_id}'";

                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = sql_update_status;
                    cmd.CommandType = CommandType.Text;
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    status = cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }

            return status;
        }

        public DataTable getFiscalYearData()
        {
            DataTable dt_fiscal_year = new DataTable();

            try
            {
                string sql_fiscal_year = "select * from fiscal_year where status = '1'";
                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = sql_fiscal_year;
                    cmd.CommandType = CommandType.Text;
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    SqlDataAdapter oda = new SqlDataAdapter(cmd);
                    oda.Fill(dt_fiscal_year);
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation(ex.Message);
            }

            return dt_fiscal_year;
        }

        public List<FiscalYear> getFiscalYearDDLList()
        {
            List<FiscalYear> fiscalYears = new List<FiscalYear>();

            DataTable dt_fin_year = getFiscalYearData();

            fiscalYears.Insert(0, new FiscalYear { fin_year_id = "0", fin_year_text = "Select" });
            foreach (DataRow dr in dt_fin_year.Rows)
            {
                string fin_year = $"{dr["fy_start_year"]}-{dr["fy_end_year"]}";

                var fiscalYear = new FiscalYear();
                fiscalYear.fin_year_id = fin_year;
                fiscalYear.fin_year_text = fin_year;

                fiscalYears.Add(fiscalYear);
            }
            return fiscalYears;
        }

        public DataTable GetClientMastersData()
        {
            DataTable dt_client_master = new DataTable();

            try
            {
                string sql_client_master = "select c.id,c.pbrsid,c.name from client  c inner join FinanceAuto_MailingList ml on ml.PbrsId = c.pbrsid where status = 1 order by name";
                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = sql_client_master;
                    cmd.CommandType = CommandType.Text;
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    SqlDataAdapter oda = new SqlDataAdapter(cmd);
                    oda.Fill(dt_client_master);
                }
            }
            catch (Exception ex)
            {
                _logger.LogInformation(ex.Message);
            }
            return dt_client_master;
        }

        public List<ClientMaster> getClientMasterDDLList()
        {
            List<ClientMaster> clientMasters = new List<ClientMaster>();

            DataTable dt_fin_year = GetClientMastersData();

            clientMasters.Insert(0, new ClientMaster { ClientID = "0", ClientNames = "Select" });
            foreach (DataRow dr in dt_fin_year.Rows)
            {
                var clientMaster = new ClientMaster();
                clientMaster.ClientID = dr["pbrsid"].ToString();
                clientMaster.ClientNames = $"{dr["name"].ToString().Replace("-", " ")}-{dr["pbrsid"]}";

                clientMasters.Add(clientMaster);
            }
            return clientMasters;
        }

        public List<MonthYear> getMonthYearDDLList()
        {
            List<MonthYear> monthYears = new List<MonthYear>();

            monthYears.Insert(0, new MonthYear { monthID = "0", month = "Select" });
            monthYears.Insert(1, new MonthYear { monthID = "Apr", month = "Apr" });
            monthYears.Insert(2, new MonthYear { monthID = "May", month = "May" });
            monthYears.Insert(3, new MonthYear { monthID = "Jun", month = "Jun" });
            monthYears.Insert(4, new MonthYear { monthID = "Jul", month = "Jul" });
            monthYears.Insert(5, new MonthYear { monthID = "Aug", month = "Aug" });
            monthYears.Insert(6, new MonthYear { monthID = "Sep", month = "Sep" });
            monthYears.Insert(7, new MonthYear { monthID = "Oct", month = "Oct" });
            monthYears.Insert(8, new MonthYear { monthID = "Nov", month = "Nov" });
            monthYears.Insert(9, new MonthYear { monthID = "Dec", month = "Dec" });
            monthYears.Insert(10, new MonthYear { monthID = "Jan", month = "Jan" });
            monthYears.Insert(11, new MonthYear { monthID = "Feb", month = "Feb" });
            monthYears.Insert(12, new MonthYear { monthID = "Mar", month = "Mar" });

            return monthYears;
        }

        public bool uploadInvoiceFile(IFormFile file, string folderPath)
        {
            var fiscal_year = _httpContextAccessor.HttpContext.Request.Form["FiscalYear.fin_year_text"].ToString();
            var client_value = _httpContextAccessor.HttpContext.Request.Form["ClientMaster.ClientNames"].ToString();
            var month_year = _httpContextAccessor.HttpContext.Request.Form["monthYear.month"].ToString();

            try
            {
                if (fiscal_year.ToLower().Equals("select") || client_value.ToLower().Equals("select") || month_year.ToLower().Equals("select") || file == null)
                {
                    return false;
                }

                /**/

                var client_name = client_value.Split('-')[0].ToString().Trim();
                var client_pbrs_id = client_value.Split('-')[1].ToString().Trim();

                /*---------- Upload And Save The File in Directory-----------*/

                string path = Path.Combine(folderPath, $"FY{fiscal_year}//{client_name}//{month_year}");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string fileName = Path.GetFileName(file.FileName);
                using (FileStream stream = new FileStream(Path.Combine(path, fileName), FileMode.Create))
                {
                    file.CopyTo(stream);
                }


                /*------- Save File Data into DB----------*/
                string insert_sql = $"insert into timesheet(pbrs_id,fy_year_id,month_year_id,file_name) values('{client_pbrs_id}','{fiscal_year}','{month_year}','{fileName.Replace("'","")}') ";

                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = insert_sql;
                    cmd.CommandType = CommandType.Text;
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message);
            }

            return true;
        }
    }
}
