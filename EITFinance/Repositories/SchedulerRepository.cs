using EITFinance.Models;
using EITFinance.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace EITFinance.Repositories
{
    public class SchedulerRepository : ISchedulerRepository
    {
        static SqlConnection conn = null;
        ICollectionSummaryService _collectionSummary;
        ISMTPService _sMTPService;
        IConfiguration _configuration;
        private readonly ILogger<SchedulerRepository> _logger;
        public SchedulerRepository(ICollectionSummaryService collectionSummary
                                 , ISMTPService sMTPService
                                 , IConfiguration configuration
                                 , ILogger<SchedulerRepository> logger)
        {
            _collectionSummary = collectionSummary;
            _sMTPService = sMTPService;
            _configuration = configuration;
            _logger = logger;
        }
        public void mailAdvicePendingClients()
        {
            try
            {
                var schedulerFlag = _configuration.GetValue<bool>("Scheduler");
                var startTime = _configuration.GetValue<string>("SchedulerStartTimespan").Split(":");
                var endTime = _configuration.GetValue<string>("SchedulerEndTimespan").Split(":");

                if (schedulerFlag == true)
                {
                    string clientName = string.Empty;
                    string remittanceStatus = _configuration.GetValue<string>("RemittanceStatus");
                    TimeSpan startAutoMail = new TimeSpan(int.Parse(startTime[0]), int.Parse(startTime[1]), int.Parse(startTime[2]));
                    TimeSpan endAutoMail = new TimeSpan(int.Parse(endTime[0]), int.Parse(endTime[1]), int.Parse(endTime[2]));
                    TimeSpan now = DateTime.Now.TimeOfDay;
                    string[] mailTo = { };
                    string[] ccTo = { };
                    var cc = _configuration.GetValue<string>("MailSettings:cc").Split(";");

                    /*----- 0. Log Info Into File -----*/
                    _logger.LogInformation("AutoMail Started");

                    /*----- 0.1. Check Today If contains In Scheduled Days In DB -----*/
                    var dayCheckFlag = scheduledDaysCheck();
                    if (dayCheckFlag)
                    {
                        /*----- 0.1. Log Info Into File -----*/
                        _logger.LogInformation("Days Check Passed");

                        /*----- 1. Check Time Range To Start Scheduler -----*/
                        if ((now >= startAutoMail) && (now <= endAutoMail))
                        {
                            /*----- 0.2. Log Info Into File -----*/
                            _logger.LogInformation("Day Timmings Check Passed");

                            /*------ 2. Get Company Names Based On Remittance Status------*/
                            JArray companyNames = _collectionSummary.getCompanyNames(remittanceStatus);

                            /*----- 3. Loop Through To get All Fin Details Of Each Company -----*/
                            if (companyNames != null)
                            {
                                /*----- 0.3. Log Info Into File -----*/
                                _logger.LogInformation("Fetching Company Names Passed");

                                foreach (JObject jobj in companyNames)
                                {
                                    clientName = jobj["client_name"].ToString();

                                    /*-- Init CC List With Values--*/
                                    List<string> ccToList = new List<string>();
                                    foreach (string _cc in cc)
                                    {
                                        ccToList.Add(_cc);
                                    }

                                    /*---- 4. Get All Pending Rows Of Comapny as Object-----*/
                                    JArray collectionSummaryOfClients = _collectionSummary.getCollectionSummaryDetails(clientName, remittanceStatus);

                                    /*----- 0.4. Log Info Into File -----*/
                                    _logger.LogInformation(clientName + " Rows Count - " + collectionSummaryOfClients.Count());

                                    foreach (JObject JObjColSum in collectionSummaryOfClients)
                                    {
                                        if (JObjColSum["to_send"].ToString() != "-" && JObjColSum["to_send"] != null && JObjColSum["to_send"].ToString() != "")
                                        {
                                            /*---- 5. Break Both Columns' String Value as Array To Get Particular Mail IDs'-----*/
                                            mailTo = JObjColSum["to_send"].ToString().Split(';');
                                            ccTo = JObjColSum["cc"].ToString().Split(';');

                                            /*---- 5.1 Add Values From ccTo Array To ccTo List-----*/
                                            foreach (string ccVal in ccTo)
                                            {
                                                ccToList.Add(ccVal);
                                            }
                                        }
                                        else
                                        {
                                            mailTo = new string[] { };
                                            ccTo = new string[] { };
                                        }
                                    }

                                    /*- Array Empty Check -*/
                                    /*----- 6. Check If there are sender names in array -----*/
                                    if (mailTo.Length > 0)
                                    {
                                        CollectionSummaryMailData colData = new CollectionSummaryMailData();
                                        colData.mailTo = mailTo;
                                        colData.CCTo = ccToList.Distinct().ToArray();
                                        colData.clientName = clientName;
                                        colData.collectionData = collectionSummaryOfClients;


                                        /*----- 7. Send Mail -----*/
                                        try
                                        {
                                            _sMTPService.sendMail(colData);
                                        }
                                        catch (Exception exinner)
                                        {
                                            /*----- Log Exception Info Into File -----*/
                                            _logger.LogInformation("SendMail Method -------- " + exinner.Message + exinner.StackTrace);
                                        }
                                    }


                                    /*----- 8. Log Info Into File -----*/
                                    _logger.LogInformation("AutoMail Finished ---- " + clientName);
                                }
                            }
                        }
                    }

                    /*----- 0. Log Info Into File -----*/
                    _logger.LogInformation("AutoMail Completed");
                }
            }
            catch (Exception ex)
            {
                /*----- Log Exception Info Into File -----*/
                _logger.LogInformation(ex.StackTrace + "--------------" + ex.Message);
            }
        }

        public bool scheduledDaysCheck()
        {
            bool daysCheckStatus = false;
            string today = System.DateTime.Today.DayOfWeek.ToString().ToLower();
            DataTable dt_scheduledDays = new DataTable();
            try
            {
                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "select item_value from site_config where item_name='scheduled_day'";
                    cmd.CommandType = CommandType.Text;
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    SqlDataAdapter oda = new SqlDataAdapter(cmd);
                    oda.Fill(dt_scheduledDays);
                }

                /* Check If Today Contains In Scheduled Day */
                if (dt_scheduledDays.Rows.Count > 0)
                {
                    string[] scheduledDays = dt_scheduledDays.Rows[0]["item_value"].ToString().Split(",");
                    if (scheduledDays.Contains(today))
                    {
                        daysCheckStatus = true;
                    }
                    else
                    {
                        daysCheckStatus = false;
                    }
                }
                else
                {
                    daysCheckStatus = false;
                }
            }
            catch (Exception ex)
            {
                /*----- 0. Log Info Into File -----*/
                _logger.LogInformation("scheduledDaysCheck ---------- " + ex.Message);
                daysCheckStatus = false;
            }
            return daysCheckStatus;
        }
    }
}
