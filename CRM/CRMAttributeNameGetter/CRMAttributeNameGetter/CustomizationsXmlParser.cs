using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;

namespace CRMAttributeNameGetter
{
    public class CustomizationsXmlParser
    {
        private List<EntityMeta> _EntityMetadatas;

        public XElement RootXml { get; set; }


        public CustomizationsXmlParser(string sourceFilePath, List<EntityMeta> entityMetadatas)
        {
            RootXml = XElement.Load(sourceFilePath);
            _EntityMetadatas = entityMetadatas;
        }

        public void GetLabels()
        {
            foreach (var entXml in RootXml.Descendants("Entity"))
            {
                // Get the entity name
                var entityName = entXml.Element("Name").Value.ToLower();

                // Get the Main Form
                var form = (
                            from fXml in entXml.Elements("FormXml")
                            from f in fXml.Elements("forms")
                            where f.Attribute("type").Value == "main"
                            select f
                            ).FirstOrDefault();

                // Prepare the specific entity in the EntityGetter
                var filteredEntityGetter = _EntityMetadatas
                    .FirstOrDefault(x => x.LogicalName.ToLower() == entityName.ToLower());

                if (form != null && filteredEntityGetter != null)
                    PopulateLabels(filteredEntityGetter, new XElement(form));
            }
        }

        private void PopulateLabels(EntityMeta currentEntity, XElement form)
        {
            var matches = from cell in (IEnumerable<Object>)form.XPathEvaluate("//row/cell")
                          where (cell as XElement).Elements("control").Any()
                            && (cell as XElement).Descendants("label").Any()
                          select new
                          {
                              LogicalName = (cell as XElement).Element("control").Attribute("id").Value,
                              Label = (cell as XElement).Element("labels").Element("label").Attribute("description").Value
                          };

            foreach (var match in matches)
            {
                // Put the label into the Attribute that matches based on LogicalName
                // and is in the Active solution, or in System solution
                if (currentEntity.Attributes
                        .Any(x => x.LogicalName == match.LogicalName && x.SolutionUniqueName.ToLower() == "active"))
                    currentEntity.Attributes
                        .First(x => x.LogicalName == match.LogicalName && x.SolutionUniqueName.ToLower() == "active")
                        .Label = match.Label;
                else if (currentEntity.Attributes.Any(x => x.LogicalName == match.LogicalName))
                    currentEntity.Attributes
                        .First(x => x.LogicalName == match.LogicalName)
                        .Label = match.Label;
            }
        }

        //public string PrintAttributes(Dictionary<string, string> attributes)
        //{
        //    var sb = new StringBuilder();
        //    var maxAttLength = attributes.Max(x => x.Key.Length);
        //    var maxLength = maxAttLength + EntityName.Length + 1;
        //    var format = "{0,-" + maxLength + "}  ->  {1}";

        //    foreach (var att in attributes)
        //    {
        //        sb.AppendLine(String.Format(format, EntityName + "." + att.Key, att.Value));
        //    }
        //    sb.AppendLine("----------------------------------------------------------");
        //    sb.AppendLine("----------------------------------------------------------");
        //    sb.AppendLine(String.Format("Done displaying attributes for entity {0}", EntityName));
        //    sb.AppendLine("----------------------------------------------------------");
        //    sb.AppendLine("----------------------------------------------------------");

        //    return sb.ToString();
        //}
    }

    public class LabelMeta
    {
        public string LogicalName { get; set; }
        public string Label { get; set; }
        public string EntityName { get; set; }
    }
}
