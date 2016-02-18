using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using Excel = Microsoft.Office.Interop.Excel;

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
            if (!FileIo.IsUConfigExist)
            {
                UserDetails.GetUserInfo();
                FileIo.ExportUserConfig();
            }
            else
            {
                FileIo.ImportUserConfig();
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
