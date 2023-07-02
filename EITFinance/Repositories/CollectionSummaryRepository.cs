using EITFinance.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace EITFinance.Repositories

{
    public class CollectionSummaryRepository : ICollectionSummaryService
    {
        IConfiguration _configuration;
        static SqlConnection conn = null;
        private readonly ILogger<SchedulerRepository> _logger;
        public CollectionSummaryRepository(IConfiguration configuration, ILogger<SchedulerRepository> logger)
        {
            _configuration = configuration;
            _logger = logger;
        }

        public dynamic getCompanyNames(string remittanceType)
        {
            dynamic obj = new JObject();
            dynamic objArray = new JArray();
            DataTable dt = new DataTable();

            string sp_name = "get_company_name";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_name;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@remittance_type", remittanceType);
                if (conn.State != ConnectionState.Open)
                {
                    conn.Open();
                }
                SqlDataAdapter oda = new SqlDataAdapter(cmd);
                oda.Fill(dt);
            }

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    obj = new JObject();
                    obj.client_name = dr["client_name"].ToString();
                    objArray.Add(obj);
                }
            }

            return objArray;
        }

        public dynamic getCollectionSummaryDetails(string clientName, string remittanceStatus)
        {
            dynamic obj_coll_summ = new JObject();
            dynamic objArray_coll_summ = new JArray();
            DataTable dt_coll_summ = new DataTable();

            try
            {

                using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
                {
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "get_collection_details";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@client_name", clientName);
                    cmd.Parameters.AddWithValue("@remittance_status", remittanceStatus);
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    SqlDataAdapter oda = new SqlDataAdapter(cmd);
                    oda.Fill(dt_coll_summ);
                }

                if (dt_coll_summ.Rows.Count > 0)
                {
                    foreach (DataRow dr_coll_summ in dt_coll_summ.Rows)
                    {
                        obj_coll_summ = new JObject();
                        obj_coll_summ.client_name = dr_coll_summ["client_name"].ToString();
                        obj_coll_summ.cheque_date = dr_coll_summ["cheque_date"].ToString();
                        obj_coll_summ.cheque_details = dr_coll_summ["cheque_details"].ToString();
                        obj_coll_summ.payment_received_date = dr_coll_summ["payment_received_date"];
                        obj_coll_summ.amount = double.Parse(dr_coll_summ["amount"].ToString()).ToString("0.00");
                        obj_coll_summ.currency = dr_coll_summ["currency"].ToString();
                        obj_coll_summ.payment_type = dr_coll_summ["payment_type"].ToString();
                        obj_coll_summ.payment_details = dr_coll_summ["payment_details"].ToString();
                        obj_coll_summ.remittance_status = dr_coll_summ["remittance_status"].ToString();
                        obj_coll_summ.to_send = dr_coll_summ["to_send"].ToString();
                        obj_coll_summ.cc = dr_coll_summ["cc"].ToString();
                        obj_coll_summ.mail_remarks = dr_coll_summ["mail_remarks"].ToString();
                        objArray_coll_summ.Add(obj_coll_summ);
                    }

                    _logger.LogInformation("getCollectionSummaryDetails --- Rows Count - " + dt_coll_summ.Rows.Count);
                }
            }
            catch (Exception ex)
            {
                obj_coll_summ = new JObject();
                obj_coll_summ.Error = ex.Message;
                objArray_coll_summ.Add(obj_coll_summ);

                _logger.LogInformation("getCollectionSummaryDetails ----------- " + ex.Message);
            }

            return objArray_coll_summ;
        }

        public int truncateCollectionSummaryStagingTable()
        {
            int status = 0;

            string sql_truncate_cc_sum__staging_tbl = "truncate table collection_summary_staging";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_truncate_cc_sum__staging_tbl;
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

        public int truncateCollectionSummaryTable()
        {
            int status = 0;

            string sql_truncate_cc_sum_tbl = "truncate table collection_summary";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql_truncate_cc_sum_tbl;
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

        public int mergeFromStagingToMainTable()
        {
            int status = 0;

            string sp_merge = "Merge_to_collection_summary";

            using (conn = new SqlConnection(_configuration.GetConnectionString("connEITFINDB")))
            {
                SqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sp_merge;
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
