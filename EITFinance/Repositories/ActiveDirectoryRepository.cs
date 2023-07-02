using EITFinance.Services;
using Microsoft.AspNetCore.Authentication;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.Protocols;
using System.Net;

namespace EITFinance.Repositories
{
    public class ActiveDirectoryRepository : IActiveDirectoryService
    {
        string LDAPServer;
        IConfiguration _configuration;
        private readonly ILogger<SchedulerRepository> _logger;
        public ActiveDirectoryRepository(IConfiguration configuration, ILogger<SchedulerRepository> logger)
        {
            _configuration = configuration;
            _logger = logger;
            LDAPServer = _configuration.GetValue<string>("LDAPServer");
        }

        public dynamic Authenticate(string username, string password)
        {
            dynamic objAuthenticate = new JObject();
            try
            {
                LdapConnection connection = new LdapConnection(LDAPServer);
                NetworkCredential credential = new NetworkCredential(username, password);
                connection.Credential = credential;
                connection.Bind();

                objAuthenticate.status = true;
                objAuthenticate.message = "success";
            }
            catch (LdapException lexc)
            {
                objAuthenticate.status = true;
                objAuthenticate.message = lexc.Message;

                _logger.LogInformation("AD Authenticate ---- " + lexc.Message);
            }
            catch (Exception ex)
            {
                objAuthenticate.status = false;
                objAuthenticate.message = ex.Message;

                _logger.LogInformation("AD Authenticate ---- " + ex.Message);
            }

            return objAuthenticate;
        }

        public dynamic getUsernameDetails(string username)
        {
            dynamic usernameDetails = new JObject();
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, LDAPServer))
                {
                    var usr = UserPrincipal.FindByIdentity(context, username);
                    if (usr != null)
                    {
                        usernameDetails.status = true;
                        usernameDetails.Name = usr.DisplayName;
                    }
                }
            }
            catch (LdapException lexc)
            {
                usernameDetails.status = false;
                usernameDetails.Message = lexc.ToString();

                _logger.LogInformation("AD Get Username Details ---- " + lexc.Message);
            }
            catch (Exception ex)
            {
                usernameDetails.status = false;
                usernameDetails.Message = ex.StackTrace;

                _logger.LogInformation("AD Get Username Details ---- " + ex.Message);
            }

            return usernameDetails;
        }
    }
}
