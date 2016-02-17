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
    static class Program
    {
        /// <summary>
        /// Stores all of the applications or items within the sharepoint list
        /// </summary>
        internal static readonly List<AppDetail> Applications = new List<AppDetail>();

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
        static void Main(string[] args)
        {
            if (!FileIo.IsUConfigExist)
            {
                UserDetails.GetUserInfo();
                UserDetails.SaveUserConfig();
            }
            else
            {
                FileIo.ImportUserConfig();
            }

            while (true)
            {
                Console.Clear();
                Console.WriteLine("Welcome to the main menu " + Username);
                Console.WriteLine("1. Extract from SharePoint list");
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
