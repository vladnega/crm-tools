using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CRMAttributeNameGetter
{
    public class SolutionManager
    {

        private Guid DefaultPublisherId = new Guid("{d21aab71-79e7-11dd-8874-00188b01e34f}");
        private const string SolutionUniqueName = "EntityAttributeComparer";
        private const string SolutionFriendlyname = "Entity Attribute Comparer";

        private CrmObject Crm;
        private Dictionary<Guid, string> Entities;

        public SolutionManager(CrmObject Crm, Dictionary<Guid, string> entities)
        {
            this.Crm = Crm;
            this.Entities = new Dictionary<Guid, string>(entities);
        }

        public byte[] ExportSolution()
        {
                CreateComparerSolution(Entities
                    .Select(x => x.Value)
                    .Aggregate((acc, val) => acc += ", " + val));

            foreach(var entity in Entities)
                AddEntityToSolution(entity);

            return GetSolutionFile();
        }

        private byte[] GetSolutionFile()
        {
            var req = new ExportSolutionRequest()
            {
                Managed = false,
                SolutionName = SolutionUniqueName
            };

            var resp = (ExportSolutionResponse)Crm.Service.Execute(req);

            return resp.ExportSolutionFile;
        }

        private void AddEntityToSolution(KeyValuePair<Guid, string> entity)
        {
            // Get the required entity
            var entMetadataId = entity.Key;
            if(entMetadataId == null)
                throw new Exception("Could not retrieve entity metadata for " + entity.Value);

            
            // Add the entity to the solution
            var addReq = new AddSolutionComponentRequest()
            {
                ComponentType = 1,  // Entity
                ComponentId = (Guid)entMetadataId,
                SolutionUniqueName = SolutionUniqueName
            };
            Crm.Service.Execute(addReq);
        }

        private void CreateComparerSolution(string entityName)
        {
            var comparerSol = new Entity("solution");

            // Check if it already exists
            var existSol = (from s in Crm.OrgContext.CreateQuery("solution")
                            where (string)s["uniquename"] == SolutionUniqueName
                            select (Guid)s["solutionid"]).ToList();

            // If it does exist, then delete it
            if (existSol.Count > 0)
                Crm.Service.Delete("solution", existSol[0]);

            // Create a new solution
            comparerSol["uniquename"] = SolutionUniqueName;
            comparerSol["friendlyname"] = SolutionFriendlyname;
            comparerSol["publisherid"] = new EntityReference("publisher", DefaultPublisherId);
            comparerSol["description"] = entityName;
            comparerSol["version"] = "1.0";
            comparerSol.Id = Crm.Service.Create(comparerSol);
        }

    }
}
