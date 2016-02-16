using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExtracFromSharepoint
{
    static class Program
    {
        internal static List<AppDetail> Applications = new List<AppDetail>();

        internal static string Password = "";

        internal static string Domain = "";

        internal static string Username = "";

        static void Main(string[] args)
        {
            ConsoleKeyInfo key;

            Console.Write("Enter password: ");
            //do
            //{
            //    key = Console.ReadKey(true);

            //    // Ignore any key out of range.
            //    //if (((int)key.Key) >= 65 && ((int)key.Key <= 90))
            //    //{
            //        // Append the character to the password.
            //        SecurePwd.AppendChar(key.KeyChar);
            //        Console.Write("*");
            //    //}
            //    // Exit if Enter key is pressed.
            //} while (key.Key != ConsoleKey.Enter);
            Password = Console.ReadLine();
            Console.Clear();
            Console.WriteLine("Enter website Domain");
            Domain = Console.ReadLine();
            Console.WriteLine("Please enter your username");
            Username = Console.ReadLine();
            

            // Create a new IE driver and navigate to the url
            var ieDriver = new InternetExplorerDriver();
            ieDriver.Navigate().GoToUrl("");
            // Wait for the user to log in and go through security concerns
            Console.WriteLine("Please log in to the sharepoint site and wait for it to load, once complete please press enter");
            Console.ReadLine();
            
            var mainStopwatch = new Stopwatch();
            mainStopwatch.Start();
            Console.WriteLine("Looking at the site");
            // Retreive all of the rows
            var all = ieDriver.FindElements(By.ClassName("ms-itmhover"));
            
            Console.WriteLine("Found " + all.Count + " applications in " + mainStopwatch.ElapsedMilliseconds + " milliseconds");
            mainStopwatch.Restart();
            // Grab all of the links to the apps
            var allLinks = new string[all.Count];
            var i = 0;
            foreach (IWebElement element in all)
            {
                allLinks[i] = element.FindElement(By.ClassName("ms-vb-title")).FindElement(By.TagName("a")).GetAttribute("href");
                i++;
                Console.WriteLine(i + "/" + all.Count + " Item urls discovered");
            }
            Console.WriteLine("All items urls discovered in " + mainStopwatch.ElapsedMilliseconds + " milliseconds");
            mainStopwatch.Restart();
            GetPageDetails(allLinks, ieDriver, mainStopwatch);
            DownloadObjects();
            SaveToExcel();
            Console.ReadLine();
        }

        static void GetPageDetails(string[] allLinks, InternetExplorerDriver ieDriver, Stopwatch sw)
        {
            for (var i = 0; i < allLinks.Length; i++)
            {
                Console.WriteLine("Checking item no " + (i+1) + "/" + allLinks.Length);
                ieDriver.Navigate().GoToUrl(allLinks[i]);
                Thread.Sleep(1000);
                var tempApp = new AppDetail(ieDriver);
                Applications.Add(tempApp);
                Console.WriteLine("Checked URL in a total of " + sw.ElapsedMilliseconds + " milliseconds, eta for rest is " + sw.ElapsedMilliseconds*(allLinks.Length-(i+1)));
                sw.Restart();
            }
            Console.WriteLine("Retreived data from all items");
        }

        static void DownloadObjects()
        {
            Console.WriteLine("Starting download on all files");
            Console.WriteLine("First collecting the urls to download from");
            var urls = new List<string>();
            var fileNames = new List<string>();

            foreach (var app in Applications)
            {
                foreach (var property in app.Properties)
                {
                    if (property.Contains("||(|)||"))
                    {
                        var split = property.Split(new[] {"||(|)||"}, StringSplitOptions.None);
                        for (var k = 1; k < split.Length; k=k+2)
                        {
                            var details = split[k].Split(new [] { "||()||"}, StringSplitOptions.None);
                            urls.Add(details[0]);
                            fileNames.Add(details[1]);
                        }
                    }
                }
            }

            Console.WriteLine("Found " + urls.Count + " files to download");
            
            using (var client = new WebClient())
            {
                client.Credentials = new NetworkCredential(Username,Password,Domain);
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                for (var i = 0; i < urls.Count; i++)
                {
                    client.DownloadFile(urls[i],"Objects\\" + fileNames[i]);
                    Console.WriteLine("Downloaded file " + (i+1));
                }
            }
        }

        static void SaveToExcel()
        {
            Console.WriteLine("Starting export to excel");
            var xlApp = new Excel.Application();

            object misValue = System.Reflection.Missing.Value;

            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.Item[1];
            xlWorkSheet.Cells[1, 1] = "Sheet 1 content";

            for (var i = 0; i < Applications[0].ProperyNames.Count; i++)
            {
                xlWorkSheet.Cells[1, (i + 1)] = Applications[0].ProperyNames[i];
            }

            for (var i = 0; i < Applications.Count; i++)
            {
                for (var j = 0; j < Applications[i].Properties.Count; j++)
                {
                    xlWorkSheet.Cells[(i+1), (j+1)] = Applications[i].Properties[j];
                }
            }

            xlWorkSheet.Shapes.AddOLEObject(
                Directory.GetCurrentDirectory() +
                "\\Objects\\Apps AIS - list of apps gone through internal review 08022016.msg", 500, 500);

            xlWorkBook.SaveAs(Directory.GetCurrentDirectory() + "\\Export.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Exception Occured while releasing object " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
