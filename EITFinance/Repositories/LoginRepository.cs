using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.Data;
using System.Data.SqlClient;

namespace EITFinance.Repositories
{
    public class LoginRepository : ILoginService
    {
        IConfiguration _configuration;
        IActiveDirectoryService _ad;
        IHttpContextAccessor _httpContextAccessor;
        static SqlConnection conn = null;
        private readonly ILogger<SchedulerRepository> _logger;
        public LoginRepository(IActiveDirectoryService ad
                               , IConfiguration configuration
                               , IHttpContextAccessor httpContextAccessor
                               , ILogger<SchedulerRepository> logger)
        {
            _ad = ad;
            _logger = logger;
            _configuration = configuration;
            _httpContextAccessor = httpContextAccessor;
        }
        public bool authenticate(string username, string password)
        {
            bool authStatus = false;
            try
            {
                DataTable dtUserMaster = new DataTable();
                JObject userObj = new JObject();
                JObject authvalue = _ad.Authenticate(username, password);
                if (bool.Parse(authvalue["status"].ToString()) == true)
                {
                    string sql_userMaster = $"select * from user_master where username = '{username}'";

                    using (SqlConnection conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                    {
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = sql_userMaster;
                        cmd.CommandType = CommandType.Text;
                        if (conn.State != ConnectionState.Open)
                        {
                            conn.Open();
                        }
                        SqlDataAdapter oda = new SqlDataAdapter(cmd);
                        oda.Fill(dtUserMaster);
                    }

                    if (dtUserMaster.Rows.Count > 0)
                    {
                        CookieOptions option = new CookieOptions();
                        option.Expires = DateTime.Now.AddMinutes(30);
                        _httpContextAccessor.HttpContext.Response.Cookies.Append("Role", dtUserMaster.Rows[0]["role"].ToString(), option);
                        _httpContextAccessor.HttpContext.Session.SetString("username", username);
                        _httpContextAccessor.HttpContext.Session.SetString("role", dtUserMaster.Rows[0]["role"].ToString());
                        _httpContextAccessor.HttpContext.Session.SetString("emp_id", dtUserMaster.Rows[0]["emp_id"].ToString());

                        authStatus = true;
                    }
                    else
                    {
                        authStatus = false;
                    }
                }
            }
            catch (Exception ex)
            {
                authStatus = false;

                _logger.LogInformation("Login Authenticate ---- " + ex.Message);
            }

            return authStatus;
        }

        public bool Logout()
        {
            bool logoutStatus = false;
            try
            {
                CookieOptions option = new CookieOptions();
                option.Expires = DateTime.Now.AddDays(-1);
                option.Secure = false;
                option.IsEssential = true;
                _httpContextAccessor.HttpContext.Response.Cookies.Append("Role", string.Empty, option);
                _httpContextAccessor.HttpContext.Response.Cookies.Delete("Role");
                _httpContextAccessor.HttpContext.Session.Clear();

                logoutStatus = true;
            }
            catch (Exception ex)
            {
                logoutStatus = false;

                _logger.LogInformation("Logout ---- " + ex.Message);
            }          

            return logoutStatus;
        }

        public dynamic getAuthorizationStatus(string loginID,string module)
        {
            string sp_authorization_status = "get_authorization_status";
            DataTable dt_authorization_status = new DataTable();

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_authorization_status;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@loginID", loginID);
                cmd.Parameters.AddWithValue("@module_name", module);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlDataAdapter oda = new SqlDataAdapter(cmd);
                oda.Fill(dt_authorization_status);
            }

            return dt_authorization_status;
        }
    }
}

