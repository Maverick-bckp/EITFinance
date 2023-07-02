using EITFinance.Controllers;
using EITFinance.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace EITFinance.Worker_Service
{
    public class TimesheetScheduler : IHostedService, IDisposable
    {
        private Timer timer;
        IConfiguration _configuration;
        ITimesheetService _timesheetService;
        private readonly ILogger<TimesheetController> _logger;
        public TimesheetScheduler(ITimesheetService timesheetService, IConfiguration configuration, ILogger<TimesheetController> logger)
        {
            _timesheetService = timesheetService;
            _configuration = configuration;
            _logger = logger;
        }

        public void Dispose()
        {
            timer?.Dispose();
        }
        public Task StartAsync(CancellationToken cancellationToken)
        {            
            var schedulerGap = _configuration.GetValue<int>("Application:SchedulerInterval"); 
            //timer = new Timer(o => _timesheetService.TimesheetProcessor(), null, TimeSpan.Zero, TimeSpan.FromMinutes(schedulerGap));

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }
}