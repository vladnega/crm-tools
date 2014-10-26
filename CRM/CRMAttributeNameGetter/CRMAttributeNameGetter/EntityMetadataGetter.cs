using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Data.SqlClient;

namespace CRMAttributeNameGetter
{
    public class EntityMetadataGetter
    {
        private bool _EntitiesRetrieved = false;
        private ICollection<EntityMeta> _Entities;

        public CrmObject Crm { get; set; }
        public ICollection<EntityMeta> Entities
        {
            get
            {
                if (!_EntitiesRetrieved)
                    GetAllMetadataFromDb();
                return _Entities;
            }
        }

        public EntityMetadataGetter(CrmObject crm)
        {
            Crm = crm;
        }

        public void ShowEntityAttributes(string entityName)
        {
            if (!Entities.Any(x => x.LogicalName == entityName))
                throw new Exception("Could not find entity with name " + entityName);

            var MaxAttLength = Entities
                .First(x => x.LogicalName == entityName)
                .Attributes
                .Max(x => x.LogicalName.Length);
            var maxLength = MaxAttLength + entityName.Length + 1;
            var format = "{0,-" + maxLength + "}  ->  {1}";

            Console.WriteLine("Display names:");
            foreach (var att in Entities.First(x => x.LogicalName == entityName).Attributes)
            {
                Console.WriteLine(String.Format(format, entityName + "." + att.LogicalName, att.DisplayName));
            }
            Console.WriteLine("Done displaying for entity {0}", entityName);
        }

        public void GetAllMetadataFromDb()
        {
            _Entities = new List<EntityMeta>();
            _EntitiesRetrieved = true;

            var connectionManager = ConfigurationManager.ConnectionStrings["Everest_MSCRM"];

            using (SqlConnection connection = new SqlConnection(connectionManager.ConnectionString))
            {
                using (SqlCommand command = new SqlCommand(
                        @"  
                            SELECT e.LogicalName, e.Name, a.LogicalName, l.Label, e.EntityId, s.UniqueName	
                            FROM MetadataSchema.Attribute a
                                INNER JOIN MetadataSchema.LocalizedLabel l ON a.AttributeId = l.ObjectId
                                INNER JOIN MetadataSchema.Entity e ON a.EntityId = e.EntityId
	                            INNER JOIN Solution s ON a.SolutionId = s.SolutionId AND e.SolutionId = s.SolutionId AND l.SolutionId = s.SolutionId
                            WHERE l.ObjectColumnName = 'DisplayName'
                                AND e.IsCustomizable = 1
                                AND e.IsImportable = 1
                        ", connection))
                {
                    connection.Open();
                    var reader = command.ExecuteReader();

                    try
                    {
                        while (reader.Read())
                        {
                            var currentEntityLogicalName = (string)reader[0];

                            // Setup the AttributeMeta object
                            var attMeta = new AttributeMeta()
                            {
                                LogicalName = (string)reader[2],
                                DisplayName = (string)reader[3],
                                EntityName = (string)reader[0],
                                SolutionUniqueName = (string)reader[5]
                            };

                            // Check if the current entity already exists
                            if (_Entities.Any(x => x.LogicalName == currentEntityLogicalName))
                            {
                                _Entities.First(x => x.LogicalName == currentEntityLogicalName)
                                    .Attributes
                                    .Add(attMeta);
                            }
                            else
                            {
                                // Set the entity name
                                var toUse = new EntityMeta()
                                {
                                    LogicalName = (string)reader[0],
                                    LocalizedName = (string)reader[1],
                                    EntityId = (Guid)reader[4]
                                };
                                // Add the attribute
                                toUse.Attributes = new List<AttributeMeta>(){attMeta};

                                _Entities.Add(toUse);
                            }
                        }
                    }
                    finally
                    {
                        reader.Close();
                    }
                }
            }
        }

    }

    public class EntityMeta
    {
        public Guid EntityId { get; set; }
        public string LogicalName { get; set; }
        public string LocalizedName { get; set; }
        public ICollection<AttributeMeta> Attributes { get; set; }

        public EntityMeta()
        { }

        public EntityMeta(EntityMeta newE)
        {
            this.LocalizedName = newE.LocalizedName;
            this.LogicalName = newE.LogicalName;
            this.Attributes = new List<AttributeMeta>(newE.Attributes);
        }
    }

    public class AttributeMeta
    {
        public string LogicalName { get; set; }
        public string DisplayName { get; set; }
        public string Label { get; set; }
        public string EntityName { get; set; }
        public string SolutionUniqueName { get; set; }
        public bool IsLabelMatchingDisplayName
        {
            get
            {
                return Label == DisplayName;
            }
        }
    }
}
