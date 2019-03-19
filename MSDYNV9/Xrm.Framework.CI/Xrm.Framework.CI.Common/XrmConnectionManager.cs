using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xrm.Framework.CI.Common.Logging;

namespace Xrm.Framework.CI.Common
{
    public class XrmConnectionManager
    {
        #region Variables

        private int DefaultTime = 120;
        private TimeSpan ConnectPolingInterval = TimeSpan.FromSeconds(15);
        private int ConnectRetryCount = 3;

        #endregion

        #region Properties

        protected ILogger Logger
        {
            get;
            set;
        }

        #endregion

        #region Constructors

        public XrmConnectionManager(ILogger logger)
        {
            Logger = logger;
        }

        #endregion

        #region Methods

        public IOrganizationService Connect(
            string connectionString,
            int timeout)
        {
            SetSecurityProtocol();
            return ConnectToCRM(connectionString, timeout);
        }

        private void SetSecurityProtocol()
        {
            Logger.LogVerbose("Current Security Protocol: {0}", ServicePointManager.SecurityProtocol);

            if (!ServicePointManager.SecurityProtocol.HasFlag(SecurityProtocolType.Tls11))
            {
                ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol ^ SecurityProtocolType.Tls11;
            }
            if (!ServicePointManager.SecurityProtocol.HasFlag(SecurityProtocolType.Tls12))
            {
                ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol ^ SecurityProtocolType.Tls12;
            }

            Logger.LogVerbose("Modified Security Protocol: {0}", ServicePointManager.SecurityProtocol);
        }

        private IOrganizationService ConnectToCRM(string connectionString, int timeout)
        {
            var timeoutDuration = new System.TimeSpan(0, 0, timeout == 0 ? DefaultTime : timeout);

            // Temporary hack to identify the unofficial AuthType S2S
            var connectionStringParser = new DbConnectionStringBuilder { ConnectionString = connectionString };
            var authenticateWithSecret = "S2S".Equals(connectionStringParser["AuthType"] as string, StringComparison.InvariantCultureIgnoreCase);

            CrmServiceClient serviceClient = null;
            for (int i = 1; i <= ConnectRetryCount; i++)
            {
                Logger.LogVerbose("Connecting to CRM [attempt {0}]...", i);
                serviceClient = authenticateWithSecret ? ConnectToCRMWithSecret(connectionStringParser, timeoutDuration) : new CrmServiceClient(connectionString);

                if (serviceClient != null && serviceClient.IsReady)
                {
                    if (serviceClient.OrganizationServiceProxy != null)
                    {
                        serviceClient.OrganizationServiceProxy.Timeout = timeoutDuration;
                        serviceClient.OrganizationServiceProxy?.EnableProxyTypes(Assembly.GetAssembly(typeof(Entities.Solution)));
                    }

                    Logger.LogVerbose("Connection to CRM Established");

                    return serviceClient;
                }
                else
                {
                    Logger.LogWarning(serviceClient.LastCrmError);
                    if (serviceClient.LastCrmException != null)
                    {
                        Logger.LogWarning(serviceClient.LastCrmException.Message);
                    }
                    if (i != ConnectRetryCount)
                        Thread.Sleep(ConnectPolingInterval);
                }
            }

            throw new Exception(string.Format("Couldn't connect to CRM instance after {0} attempts: {1}", ConnectRetryCount, serviceClient?.LastCrmError));
        }

        //
        // Temporary hack to support authentication with clientid and secret - while we wait for official support from Microsoft...
        // To have minimal impact on existing code a new (unofficial) AuthType "S2S" has been introduced in the connectionstring.
        // Sample connectionstrings:
        //  => "AuthType=S2S;Url=https://organization-name.crm.dynamics.com;ClientId=abcdefgh-1234-4321-abcd-abcdefghijkl;Secret=superdupersecretthatnonewilleverguess;tenantid=abcdefgh-1234-4321-abcd-abcdefghijkl"
        //  => "AuthType=S2S;Url=https://organization-name.crm.dynamics.com;ClientId=abcdefgh-1234-4321-abcd-abcdefghijkl;Secret=superdupersecretthatnonewilleverguess;tenantid=abcdefgh-1234-4321-abcd-abcdefghijkl;serviceurl=https://organization-name.api.crm.dynamics.com/XRMServices/2011/Organization.svc/web"
        //  => "AuthType=S2S;Url=https://organization-name.crm.dynamics.com;ClientId=abcdefgh-1234-4321-abcd-abcdefghijkl;Secret=superdupersecretthatnonewilleverguess;authority=https://login.microsoftonline.com/abcdefgh-1234-4321-abcd-abcdefghijkl/"
        //
        private CrmServiceClient ConnectToCRMWithSecret(DbConnectionStringBuilder connectionStringParser, TimeSpan timeoutDuration)
        {
            var url = connectionStringParser.GetValue<string>("url", true);
            var clientId = connectionStringParser.GetValue<string>("clientId", true);
            var secret = connectionStringParser.GetValue<string>("secret", true); 
            var authority = connectionStringParser.GetValue<string>("authority", false) ?? $"https://login.microsoftonline.com/{connectionStringParser.GetValue<string>("tenantid", true)}/";
            var serviceUrl = connectionStringParser.GetValue<string>("serviceUrl", false) ?? $"{url}/XRMServices/2011/Organization.svc/web";

            Logger.LogVerbose("Connecting to '{0}' with client id '{1}'", url, clientId);

            var clientCredential = new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential(clientId, secret);

            var authenticationContext = new Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext(authority);
            var authenticationResult = authenticationContext.AcquireToken(url, clientCredential);

            var organizationWebProxyClient = new Microsoft.Xrm.Sdk.WebServiceClient.OrganizationWebProxyClient(new Uri(serviceUrl), timeoutDuration, Assembly.GetAssembly(typeof(Entities.Solution))) { HeaderToken = authenticationResult.AccessToken };
            return new CrmServiceClient(organizationWebProxyClient);
        }
        #endregion
    }

    public static class ConnectionStringParserExtensions
    {
        public static T GetValue<T>(this DbConnectionStringBuilder connectionStringParser, string key, bool throwOnNotFound = false)
        {
            if (!connectionStringParser.TryGetValue(key, out var value) && throwOnNotFound)
            {
                throw new ArgumentNullException($"{key} missing from connectionstring");
            }

            return (T)value;
        }
    }
}
