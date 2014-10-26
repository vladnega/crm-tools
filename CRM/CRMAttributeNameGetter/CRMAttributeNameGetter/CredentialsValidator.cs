using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRMAttributeNameGetter
{
    public class CredentialsValidator
    {
        private MessagePrompter Mp;

        private string Url;
        private string Domain;
        private string Username;
        private string Password;

        public CredentialsValidator(MessagePrompter mp)
        {
            Mp = mp;
        }

        public CrmObject GetCrmConnection()
        {
            // Get the Url
            Mp.Prompt("Please enter the Url of the organization. E.g. 'http://www.contoso.com:5555/OrganizationName'\n");
            Url = Mp.Read();

            // Get the Domain
            Mp.Prompt("Please enter the Domain of the user.");
            Domain = Mp.Read();

            // Get the Username
            Mp.Prompt("Please enter the Username");
            Username = Mp.Read();

            // Get the Password
            Mp.Prompt("Please enter the Password");
            Password = Mp.ReadPassword();

            Mp.Prompt("\n\nChecking connection details...\n\n");
            // Check the credentials
            try
            {
                CrmConnection connection = CrmConnection.Parse(
                    "Url=" + Url + "; Domain=" + Domain + "; Username=" + Username + "; Password=" + Password);
                var service = new OrganizationService(connection);
                // Execute a dummy request
                service.Execute(new RetrieveVersionRequest());

                return new CrmObject(service);
            }
            catch (Exception e)
            {
                Mp.Prompt("The connection details supplied are not valid. Please enter the correct credentials");
                return GetCrmConnection();
            }
        }

        internal CrmObject GetCrmConnectionFromString(string connectionString)
        {
            Mp.Prompt("Checking connection details...\n\n");
            CrmConnection connection = CrmConnection.Parse(connectionString);
            var service = new OrganizationService(connection);
            // Execute a dummy request
            service.Execute(new RetrieveVersionRequest());

            return new CrmObject(service);
        }
    }
}
