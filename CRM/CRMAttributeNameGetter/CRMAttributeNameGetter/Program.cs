using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using SharpCompress.Common;
using SharpCompress.Reader;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.AccessControl;
using System.Text;

namespace CRMAttributeNameGetter
{
    class Program
    {
        public enum ActionList
        {
            ShowNonMatchingLabels = 1,
            ShowAllAttributes = 2,
        }

        static void Main(string[] args)
        {
            MessagePrompter mp = new MessagePrompter();

            // Get the connection to CRM
            var cv = new CredentialsValidator(mp);
            var Crm = cv.GetCrmConnectionFromString(ConfigurationManager.ConnectionStrings["EverestOrg"].ConnectionString);

            // Enable debugging trace errors
            if (false)
                Trace.Listeners.Add(new ConsoleTraceListener() { Name = "debug" });


            // Get the list of entities from CRM and their attributes
            var entityGetter = new EntityMetadataGetter(Crm);
            mp.Prompt("Getting the list of entities from CRM");

            #region Identify which entity to analyse
            var entityName = GetSelectedEntity(entityGetter
                .Entities
                .OrderBy(x => x.LogicalName)
                .Select(x => x.LocalizedName)
                .ToList());

            mp.Prompt("You have selected \"{0}\"", entityName);
            #endregion

            // Get the selected entity object
            var selectedEntities = new List<EntityMeta>();
            if (entityName == "ALL")
                selectedEntities = entityGetter.Entities.ToList();
            else
                selectedEntities = entityGetter.Entities.Where(x => x.LocalizedName == entityName).ToList();

            #region Export the Solution from CRM
            //mp.Prompt("Exporting the solution from CRM");

            //var solManager = new SolutionManager(
            //    Crm,
            //    selectedEntities.ToDictionary(x => x.EntityId, x => x.LogicalName.ToLower()));
            //byte[] exportXml = solManager.ExportSolution();

            //var exportPath = @"C:\temp\EntityAttributeComparerSolution.zip";
            //File.WriteAllBytes(exportPath, exportXml);

            //mp.Prompt("Solution exported to {0}.", exportPath);
            #endregion


            #region Extract the customizations.xml from the solution archive
            var zipPath = @"C:\temp\EntityAttributeComparerSolution.zip";
            var extractPath = @"C:\temp\EntityAttributeComparerSolution";

            ExtractCustomizationsFile(zipPath, extractPath);
            #endregion


            #region Get the labels of the attributes
            //mp.Prompt("Parsing the customizations.xml to gather the attribute labels...");

            var customizationsPath = @"C:\temp\EntityAttributeComparerSolution\customizations.xml";
            var custXmlParser = new CustomizationsXmlParser(customizationsPath, selectedEntities);
            custXmlParser.GetLabels();

            //mp.Prompt("Parsing complete.");
            #endregion


            ChooseNextAction(mp, extractPath, selectedEntities);
        }

        private static void ChooseNextAction(MessagePrompter mp, string extractPath, List<EntityMeta> selectedEntities)
        {
            mp.Prompt("");
            mp.Prompt("Please choose what to do with the gathered attributes. Type in just the number.");
            mp.Prompt("[{0}] Show the attributes that have different display names and labels", (int)ActionList.ShowNonMatchingLabels);
            mp.Prompt("[{0}] Show attributes of the selected entity", (int)ActionList.ShowAllAttributes);
            mp.Prompt("[{0}] Exit", 0);

            var exit = false;
            while (!exit)
            {
                var resp = mp.Read();
                var resultsFilePath = "";

                switch (Convert.ToInt32(resp))
                {
                    case (0):
                        exit = true;
                        break;

                    case ((int)ActionList.ShowNonMatchingLabels):
                        resultsFilePath = extractPath + "\\NonMatchingLabels.out";
                        ShowAllAttributes(resultsFilePath, selectedEntities, (att) =>
                            !String.IsNullOrEmpty(att.Label)
                            && !att.IsLabelMatchingDisplayName);
                        mp.Prompt("Results are stored in the '{0}' file.", resultsFilePath);
                        break;

                    case ((int)ActionList.ShowAllAttributes):
                        resultsFilePath = extractPath + "\\AllAttributes.out";
                        ShowAllAttributes(resultsFilePath, selectedEntities, (att) =>
                            !String.IsNullOrEmpty(att.Label)
                            && !att.IsLabelMatchingDisplayName);
                        mp.Prompt("Results are stored in the '{0}' file.", resultsFilePath);
                        break;

                    default:
                        mp.Prompt("Please choose a number between {0} and {1}.", 1, 2);
                        break;
                }

                mp.Prompt("");
                mp.Prompt("Please choose the next action.");
            }
        }

        private static void ShowAllAttributes(string resultsFilePath, List<EntityMeta> selectedEntities, Func<AttributeMeta, bool> attributeFilter)
        {
            var format = "{0,-50} || {1,-40} || {2,-40} || {3,-40} || {4}";
            var sb = new StringBuilder();
            sb.AppendLine(String.Format(format,"Entity Name", "Attribute Logical name", "Attribute Display Name", "Attribute Label", "Solution"));
            sb.AppendLine(String.Format(format, "", "", "", "", ""));
            sb.AppendLine(String.Format(format, "", "", "", "", ""));

            // Match the Labels and Display Names based on logical name
            foreach (var ent in selectedEntities)
            {
                foreach (var att in ent.Attributes
                    .Where(attributeFilter)
                    .OrderBy(x => x.LogicalName))
                {
                    sb.AppendLine(String.Format(format, att.EntityName, att.LogicalName, att.DisplayName, att.Label ?? "", att.SolutionUniqueName));
                }
            }

            StreamWriter file = new StreamWriter(resultsFilePath);
            file.WriteLine(sb.ToString());
            file.Close();
        }

        private static void ExtractCustomizationsFile(string zipPath, string extractPath)
        {
            // Create directory
            Directory.CreateDirectory(extractPath);

            using (Stream stream = File.OpenRead(zipPath))
            {
                var reader = ReaderFactory.Open(stream);
                while (reader.MoveToNextEntry())
                {
                    if (reader.Entry.FilePath.Contains("customizations.xml"))
                    {
                        reader.WriteEntryToDirectory(extractPath,
                            ExtractOptions.Overwrite);
                        break;
                    }
                }
            }
        }

        private static string GetSelectedEntity(List<string> entityNames)
        {
            int entityIndex;

            // Option for getting all entities
            Console.WriteLine("[-1] ALL Entities");

            for (var i = 0; i < entityNames.Count; i++)
                Console.WriteLine("[{0}] {1}", i, entityNames[i]);

            Console.WriteLine("Please select the entity you want to inspect. Type in just the number.");

            while (true)
            {
                entityIndex = 0;
                bool isValidNumber = int.TryParse(Console.ReadLine(), out entityIndex);

                if (isValidNumber && entityIndex < entityNames.Count)
                    break;
                else
                    Console.WriteLine("Please enter a number between -1 and {0}.", entityNames.Count - 1);
            }

            if (entityIndex == -1)
                return "ALL";
            else
                return entityNames[entityIndex];
        }

        private static void ShowMessage(string format, params object[] args)
        {
            Console.WriteLine(format, args);
        }
    }
}
