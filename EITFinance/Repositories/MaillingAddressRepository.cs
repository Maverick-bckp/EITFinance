using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using EITFinance.Services;

namespace EITFinance.Repositories
{
    public class MaillingAddressRepository : IMaillingAddressService
    {
        private IConfiguration _configuration;
        IHttpContextAccessor _httpContextAccessor;
        private readonly ILogger<SchedulerRepository> _logger;
        static SqlConnection conn = null;

        public MaillingAddressRepository(IConfiguration Configuration, ILogger<SchedulerRepository> logger, IHttpContextAccessor httpContextAccessor)
        {
            _configuration = Configuration;
            _logger = logger;
            _httpContextAccessor = httpContextAccessor;
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
