using EITFinance.Repositories;
using EITFinance.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Hosting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EITFinance
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllersWithViews();
            services.AddSession();
            services.AddSingleton<IConfiguration>(Configuration);
            services.AddSingleton<ICollectionSummaryService, CollectionSummaryRepository>();
            services.AddSingleton<ISMTPService, SMTPRepository>();
            services.AddSingleton<ISchedulerRepository, SchedulerRepository>();
            services.AddSingleton<IActiveDirectoryService, ActiveDirectoryRepository>();
            services.AddSingleton<ILoginService, LoginRepository>();
            services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();
            services.AddSingleton<IBillingService, BillingRepository>();
            services.AddSingleton<IMaillingAddressService, MaillingAddressRepository>();
            services.AddSingleton<ITDSCollectionService, TDSCollectionRepository>();
            services.AddSingleton<IUnbilledRevenueService, UnbilledRevenueRepository>();
            services.AddSingleton<ITimesheetService, TimeSheetRepository>();
            services.AddSingleton<IEmailSender, EmailSender>();
            services.AddSingleton<IPOService, PORepository>();

        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
            }

            app.UseStaticFiles();
            //app.UseStaticFiles(new StaticFileOptions()
            //{
            //    FileProvider = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), @"Templates")),
            //    RequestPath = new PathString("/Templates")
            //});

            app.UseRouting();

            app.UseAuthorization();
            app.UseSession();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Login}/{action=Index}/{id?}");
            });
        }
    }
}
