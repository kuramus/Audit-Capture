using System;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;

namespace AuditCapture
{
    public class ConnectionManager
    {
        public static IOrganizationService _service;
        public static void createCRMConnection(string url_,string userName, string pass)
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = userName;
                credentials.UserName.Password = pass;                
                Uri serviceUri = new Uri(url_);
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                proxy.EnableProxyTypes();
                _service = (IOrganizationService)proxy;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}