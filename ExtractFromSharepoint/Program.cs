using System;
using System.Collections.Generic;

namespace ExtractFromSharepoint
{
    internal static class Program
    {
        /// <summary>
        /// Stores all of the applications or items within the sharepoint list
        /// </summary>
        internal static List<AppDetail> Applications = new List<AppDetail>();

        /// <summary>
        /// Password to use for active directory
        /// </summary>
        internal static string Password = "";

        /// <summary>
        /// The Active Directory domain name
        /// </summary>
        internal static string AdDomain = "";

        /// <summary>
        /// Active Directory username
        /// </summary>
        internal static string Username = "";

        /// <summary>
        /// Url to grab all the list from
        /// </summary>
        internal static string ListUrl = "";

        /// <summary>
        /// The domain of the website or the ip address(used for filtering attachments to links to other websites)
        /// </summary>
        internal static string Domian = "";

        /// <summary>
        /// Entry into the program
        /// </summary>
        /// <param name="args"></param>
        private static void Main(string[] args)
        {
            if (args[0] == "/?")
            {
                Console.WriteLine("This app currently does not support any command line arguments");
            }
            // If the user config exists then import it
            if (FileIo.IsUConfigExist)
            {
                FileIo.ImportUserConfig();
            }
            // Else ask the user for their configuration
            else
            {
                UserDetails.GetUserInfo();
                FileIo.ExportUserConfig();
            }

            while (true)
            {
                Console.Clear();
                Console.WriteLine("Welcome to the main menu " + Username);
                Console.WriteLine("1. Extract from SharePoint/file and export to excel");
                Console.WriteLine("2. Update SharePoint settings");
                Console.WriteLine("3. Exit");

                var option = Console.ReadLine();

                switch (option)
                {
                    case "1":
                        SharePointExtract.Main();
                        break;

                    case "2":
                        UserDetails.Main();
                        break;

                    case "3":
                        return;

                    default:
                        continue;
                }
            }
        }
    }
}
