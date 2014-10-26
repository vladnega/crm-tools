using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRMAttributeNameGetter
{
    public class CrmObject
    {
        private OrganizationServiceContext _OrgContext = null;

        public IOrganizationService Service { get; set; }
        public OrganizationServiceContext OrgContext
        {
            get
            {
                if (_OrgContext == null)
                    _OrgContext = new OrganizationServiceContext(Service);
                return _OrgContext;
            }
        }


        public CrmObject(IOrganizationService service)
        {
            // Make sure the received service is valid
            service.Execute(new WhoAmIRequest());
            Service = service;
        }
    }
}
