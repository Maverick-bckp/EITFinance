using EITFinance.Models;
using EITFinance.Repositories;
using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace EITFinance.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class CollectionSummaryAPIController : ControllerBase
    {
        ICollectionSummaryService _collectionSummary;
        public CollectionSummaryAPIController(ICollectionSummaryService collectionSummary)
        {
            _collectionSummary = collectionSummary;
        }


        [HttpPost("CompanyNames")]
        public dynamic getCompanyNames(string remittanceType)
        {
            var companyNames = _collectionSummary.getCompanyNames(remittanceType);
            if (companyNames == null)
            {
                return NoContent();
            }
            else
            {
                return Content(JsonConvert.SerializeObject(companyNames), "application/json");
            }
        }

        [HttpPost("CollectionSummary")]
        public dynamic getCollectionSummaryDetails(string clientName, string remittanceStatus)
        {
            var collectionSummary = _collectionSummary.getCollectionSummaryDetails(clientName, remittanceStatus);
            if (collectionSummary == null)
            {
                return NoContent();
            }
            else
            {
                return Content(JsonConvert.SerializeObject(collectionSummary), "application/json");
            }
        }
    }
}
