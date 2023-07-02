using EITFinance.Models;
using EITFinance.Models.Timesheet;
using EITFinance.Models.Timesheet.DTOs;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.Data;

namespace EITFinance.Services
{
    public interface ITimesheetService
    {
        bool UploadMaillingAddresses(IFormFile file);
        List<InsertTimesheetDTO> DirectoryScanner(string folderPath, string archiveFolderPath);
        DataTable GetTimesheets(int status);
        void Insert(List<InsertTimesheetDTO> timesheet);
        int updateStatus(string clientName);
        void TimesheetProcessor();
        List<FiscalYear> getFiscalYearDDLList();
        List<ClientMaster> getClientMasterDDLList();
        List<MonthYear> getMonthYearDDLList();
        bool uploadInvoiceFile(IFormFile file, string folderPath);
    }
}
