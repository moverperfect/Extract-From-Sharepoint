using System.IO;
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
    }
}
