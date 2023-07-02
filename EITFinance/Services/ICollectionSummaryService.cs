namespace EITFinance.Services
{
    public interface ICollectionSummaryService
    {
        public dynamic getCompanyNames(string remittanceType);
        public dynamic getCollectionSummaryDetails(string clientName, string remittanceStatus);
        int truncateCollectionSummaryTable();
        int truncateCollectionSummaryStagingTable();
        int mergeFromStagingToMainTable();
    }
}
