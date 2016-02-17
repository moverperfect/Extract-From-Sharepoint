using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Xml;

namespace ExtractFromSharepoint
{
    static class FileIo
    {
        /// <summary>
        /// True if the user config file exists
        /// </summary>
        public static bool IsUConfigExist => File.Exists("User.config");

        /// <summary>
        /// Imports the user config file into the program
        /// </summary>
        internal static void ImportUserConfig()
        {
            var r = new XmlTextReader("User.config");
            while (!r.EOF)
            {
                r.Read();
                if (r.NodeType != XmlNodeType.Element)
                {
                    continue;
                }
                switch (r.Name)
                {
                    case "AdDomain":
                        r.Read();
                        Program.AdDomain = r.Value;
                        break;

                    case "Username":
                        r.Read();
                        Program.Username = r.Value;
                        break;

                    case "ListUrl":
                        r.Read();
                        Program.ListUrl = r.Value;
                        break;

                    case "Domain":
                        r.Read();
                        Program.Domian = r.Value;
                        break;
                }
            }

            r.Close();
        }

        internal static void ExportItems(string filename)
        {
            var applications = new XmlTemplate("Applications", null, null, null, null, new List<XmlTemplate>());
            foreach (var application in Program.Applications)
            {
                var temp = new XmlTemplate("Application", null, null, null, null, new List<XmlTemplate>());
                temp.SubProperties.Add(new XmlTemplate("Items", null, null, new List<string>(), new List<string>(), new List<XmlTemplate>()));
                var property = new List<string>();
                var propertyNames = new List<string>();
                var attributes = new List<string>();
                var attributeNames = new List<string>();
                for (var i = 0; i < application.ProperyNames.Count; i++)
                {
                    property.Add(application.Properties[i]);
                    propertyNames.Add("Item");
                    attributes.Add(application.ProperyNames[i]);
                    attributeNames.Add("Name");
                }
                temp.SubProperties[0].Properties = property;
                temp.SubProperties[0].PropertyNames = propertyNames;
                temp.SubProperties[0].PropertyAttributes = attributes;
                temp.SubProperties[0].PropertyAttributeNames = attributeNames;
                applications.SubProperties.Add(temp);
            }

            var s = new XmlWriterSettings();
            s.Indent = true;
            s.IndentChars = "    ";
            var w = XmlWriter.Create(filename + ".xml", s);
            w.WriteStartDocument();
            w = applications.GetData(w);
            w.WriteEndDocument();
            w.Close();
        }
    }
}
