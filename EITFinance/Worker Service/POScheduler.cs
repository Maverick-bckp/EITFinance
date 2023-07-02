using EITFinance.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace EITFinance.Worker_Service
{
    public class POScheduler : IHostedService, IDisposable
    {
        private Timer timer;
        IConfiguration _configuration;
        IPOService _poService;
        public POScheduler(IPOService poService, IConfiguration configuration)
        {
            _poService = poService;
            _configuration = configuration;
        }

        public void Dispose()
        {
            timer?.Dispose();
        }
        public Task StartAsync(CancellationToken cancellationToken)
        {
            var schedulerGap = _configuration.GetValue<int>("SchedulerInterval"); 
            //timer = new Timer(o => _poService.ProcessPO(), null, TimeSpan.Zero, TimeSpan.FromMinutes(schedulerGap));

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }
}