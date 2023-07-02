using EITFinance.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace EITFinance.Worker_Service
{
    public class CollectionSummaryScheduler : IHostedService, IDisposable
    {
        private Timer timer;
        IConfiguration _configuration;
        ISchedulerRepository _schedulerRepository;
        public CollectionSummaryScheduler(ISchedulerRepository schedulerRepository, IConfiguration configuration)
        {
            _schedulerRepository = schedulerRepository;
            _configuration = configuration;
        }

        public void Dispose()
        {
            timer?.Dispose();
        }
        public Task StartAsync(CancellationToken cancellationToken)
        {
            var schedulerGap = _configuration.GetValue<int>("SchedulerInterval"); 
            timer = new Timer(o => _schedulerRepository.mailAdvicePendingClients(), null, TimeSpan.Zero, TimeSpan.FromMinutes(schedulerGap));

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            return Task.CompletedTask;
        }
    }
}