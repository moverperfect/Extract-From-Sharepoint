using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
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
        /// True if the excel configuration exists
        /// </summary>
        public static bool IsEConfigExist => File.Exists("Excel.config");

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

        /// <summary>
        /// Saves the users configuration to a file
        /// </summary>
        internal static void ExportUserConfig()
        {
            var w = new XmlTextWriter("User.config", Encoding.Default);

            w.WriteStartDocument();
            w.WriteStartElement("config");
            w.WriteElementString("AdDomain", Program.AdDomain);
            w.WriteElementString("Username", Program.Username);
            w.WriteElementString("ListUrl", Program.ListUrl);
            w.WriteElementString("Domain", Program.Domian);
            w.WriteEndElement();
            w.WriteEndDocument();

            w.Close();
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

            var s = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "    "
            };
            var w = XmlWriter.Create(filename + (filename.Contains(".xml")?"":".xml"), s);
            w.WriteStartDocument();
            w = applications.GetData(w);
            w.WriteEndDocument();
            w.Close();
        }

        internal static void ImportItems(string readFile)
        {
            var applications = new List<AppDetail>();
            var r = XmlReader.Create(readFile + (readFile.Contains(".xml") ? "" : ".xml"));
            var app = new AppDetail();
            while (!r.EOF)
            {
                r.Read();
                switch (r.NodeType)
                {
                    case XmlNodeType.Element:
                        switch (r.Name)
                        {
                            case "Item":
                                app.ProperyNames.Add(r.GetAttribute("Name"));
                                r.Read();
                                app.Properties.Add(r.Value);
                                break;
                        }
                        break;

                     case XmlNodeType.EndElement:
                        switch (r.Name)
                        {
                            case "Application":
                                applications.Add(app);
                                app = new AppDetail();
                                break;
                        }
                        break;
                }
            }
            r.Close();
            Program.Applications = applications;
        }

        internal static void ImportExcelConfig()
        {
            ExcelExport.Columns = new List<ExcelColumn>();
            ExcelExport.Header = new Header();
            ExcelExport.Rows = new List<Row>();

            var r = XmlReader.Create("Excel.config");
            var row = new Row();
            var header = new Header();
            var column = new ExcelColumn();
            while (!r.EOF)
            {
                r.Read();
                switch (r.NodeType)
                {
                    case XmlNodeType.Element:
                        switch (r.Name)
                        {
                            case "RowColour":
                                r.Read();
                                row.Colour = r.Value;
                                break;

                            case "RowHeight":
                                r.Read();
                                decimal height;
                                if (decimal.TryParse(r.Value, out height))
                                {
                                    row.Height = height;
                                }
                                else
                                {
                                    Console.WriteLine("ERROR READING THE VALUE " + r.Value);
                                    Console.WriteLine("Returning in 10 seconds");
                                    Thread.Sleep(10000);
                                }
                                break;

                            case "Height":
                                r.Read();
                                decimal hheight;
                                if (decimal.TryParse(r.Value, out hheight))
                                {
                                    header.Height = hheight;
                                }
                                else
                                {
                                    Console.WriteLine("ERROR READING THE VALUE " + r.Value);
                                    Console.WriteLine("Returning in 10 seconds");
                                    Thread.Sleep(10000);
                                }
                                break;

                            case "BackgroundColour":
                                r.Read();
                                header.BackgroundColour = r.Value;
                                break;

                            case "TextColour":
                                r.Read();
                                header.TextColour = r.Value;
                                break;

                            case "Name":
                                r.Read();
                                column.Name = r.Value;
                                break;

                            case "Width":
                                r.Read();
                                decimal width;
                                if (decimal.TryParse(r.Value, out width))
                                {
                                    column.Width = width;
                                }
                                else
                                {
                                    Console.WriteLine("ERROR READING THE VALUE " + r.Value);
                                    Console.WriteLine("Returning in 10 seconds");
                                    Thread.Sleep(10000);
                                }
                                break;
                        }
                        break;

                    case XmlNodeType.EndElement:
                        switch (r.Name)
                        {
                            case "Row":
                                ExcelExport.Rows.Add(row);
                                row = new Row();
                                break;

                            case "HeaderProperties":
                                ExcelExport.Header = header;
                                header = new Header();
                                break;

                            case "Item":
                                ExcelExport.Columns.Add(column);
                                column = new ExcelColumn();
                                break;
                        }
                        break;
                }
            }
            r.Close();
        }
    }
}
