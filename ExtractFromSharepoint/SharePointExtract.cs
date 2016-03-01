using System;
using System.Collections.Generic;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;

namespace ExtractFromSharepoint
{
    static class SharePointExtract
    {
        internal static void Main()
        {
            while (true)
            {
                Console.Clear();
                Console.WriteLine("There are " + Program.Applications.Count + " application/s stored");
                Console.WriteLine("Sharepoint extract menu");
                Console.WriteLine("1. Extract from sharepoint site");
                Console.WriteLine("2. Extract from file");
                Console.WriteLine("3. Extract to excel file");
                Console.WriteLine("4. Back");

                var option = Console.ReadLine();

                switch (option)
                {
                    case "1":
                        if (Program.Password == "")
                        {
                            Console.WriteLine("Your password is currently NULL");
                            Console.WriteLine("Would you like to enter a password for the user " + Program.Username + "? y/n");
                            if (Console.ReadLine() == "y")
                            {
                                UserDetails.GetUserPassword();
                            }
                        }
                        Console.WriteLine("Would you like to save the data after extraction? y/n");
                        var filename = "";
                        if (Console.ReadLine()?.ToLower() == "y")
                        {
                            Console.WriteLine("Please enter the name of the file to save to");
                            filename = Console.ReadLine() ?? "Export";
                            var unspupportedRegex = new Regex("(^(PRN|AUX|NUL|CON|COM[1-9]|LPT[1-9]|(\\.+)$)(\\..*)?$)|(([\\x00-\\x1f\\\\?*:\";‌​|/<>‌​])+)|([\\. ]+)", RegexOptions.IgnoreCase);
                            while (unspupportedRegex.IsMatch(filename))
                            {
                                Console.WriteLine("This filename is invalid, please try again");
                                filename = Console.ReadLine() ?? "Export";
                            }
                        }
                        GetAllSharePointData();
                        if (filename != "")
                        {
                            FileIo.ExportItems(filename);
                        }
                        break;

                    case "2":
                        Console.WriteLine("Please enter the name of the file to load from");
                        var readFile = Console.ReadLine();
                        FileIo.ImportItems(readFile);
                        break;

                    case "3":
                        ExcelExport.Main();
                        break;

                    case "4":
                        return;

                    default:
                        continue;
                }
            }
        }

        private static void GetAllSharePointData()
        {
            // Create a new IE driver and navigate to the url
            var ieDriver = new InternetExplorerDriver();
            ieDriver.Navigate().GoToUrl(Program.ListUrl);

            // Wait for the user to log in and go through security concerns
            Console.WriteLine("Please log in to the sharepoint site and wait for it to load, once complete please press enter");
            Console.ReadLine();

            // Grab all of the links from the list
            var allLinks = GetLinksFromSharePoint(ieDriver);

            // Get all of the details for each item
            GetPageDetails(allLinks, ieDriver);

            // Download all of the attachments
            DownloadObjects();
        }

        private static string[] GetLinksFromSharePoint(ISearchContext ieDriver)
        {
            Console.WriteLine("Looking at the site");
            // Retreive all of the rows
            var all = ieDriver.FindElements(By.ClassName("ms-itmhover"));

            Console.WriteLine("Found " + all.Count + " applications");
            // Grab all of the links to the apps
            var allLinks = new string[all.Count];
            var i = 0;
            foreach (IWebElement element in all)
            {
                allLinks[i] = element.FindElement(By.ClassName("ms-vb-title")).FindElement(By.TagName("a")).GetAttribute("href");
                i++;
                Console.WriteLine(i + "/" + all.Count + " Item urls discovered");
            }
            Console.WriteLine("All items urls discovered");
            return allLinks;
        }

        static void GetPageDetails(string[] allLinks, InternetExplorerDriver ieDriver)
        {
            for (var i = 0; i < allLinks.Length; i++)
            {
                Console.WriteLine("Checking item no " + (i + 1) + "/" + allLinks.Length);
                ieDriver.Navigate().GoToUrl(allLinks[i]);
                Thread.Sleep(1000);
                var tempApp = new AppDetail(ieDriver);
                Program.Applications.Add(tempApp);
            }
            Console.WriteLine("Retreived data from all items");
        }

        static void DownloadObjects()
        {
            Console.WriteLine("Starting download on all files");
            Console.WriteLine("First collecting the urls to download from");
            var urls = new List<string>();
            var fileNames = new List<string>();

            foreach (var app in Program.Applications)
            {
                foreach (var property in app.Properties)
                {
                    if (property.Contains("||(|)||"))
                    {
                        var split = property.Split(new[] { "||(|)||" }, StringSplitOptions.None);
                        for (var k = 1; k < split.Length; k = k + 2)
                        {
                            var details = split[k].Split(new[] { "||()||" }, StringSplitOptions.None);
                            urls.Add(details[0]);
                            fileNames.Add(details[1]);
                        }
                    }
                }
            }

            Console.WriteLine("Found " + urls.Count + " files to download");

            using (var client = new WebClient())
            {
                client.Credentials = new NetworkCredential(Program.Username, Program.Password, Program.AdDomain);
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                for (var i = 0; i < urls.Count; i++)
                {
                    client.DownloadFile(urls[i], "Objects\\" + fileNames[i]);
                    Console.WriteLine("Downloaded file " + (i + 1));
                }
            }
        }
    }
}
