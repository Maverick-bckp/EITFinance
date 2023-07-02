using EITFinance.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EITFinance.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class SchedulerController : ControllerBase
    {
        ISchedulerRepository _schedulerRepository;
        public SchedulerController(ISchedulerRepository schedulerRepository)
        {
            _schedulerRepository = schedulerRepository;
        }

        [HttpPost("MailAdvicePendingClients")]
        public void mailAdvicePendingClients()
        {
            _schedulerRepository.mailAdvicePendingClients();
        }
    }
}
