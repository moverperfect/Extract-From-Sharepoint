using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ExtractFromSharepoint
{
    /// <summary>
    /// Handles the export to excel functions
    /// </summary>
    internal static class ExcelExport
    {
        /// <summary>
        /// List of columns that will be exported to excel
        /// </summary>
        internal static List<ExcelColumn> Columns = new List<ExcelColumn>();

        /// <summary>
        /// Represents the header row that will get exported
        /// </summary>
        internal static Header Header = new Header();

        /// <summary>
        /// Represents the rows that will be exported
        /// </summary>
        internal static List<Row> Rows = new List<Row>();

        /// <summary>
        /// Entry point into the excel export
        /// </summary>
        internal static void Main()
        {
            // Import the excel config if it exists
            if (FileIo.IsEConfigExist) FileIo.ImportExcelConfig();

            while (true)
            {
                Console.Clear();
                Console.WriteLine("Excel Export Main Menu");
                Console.WriteLine("Config loaded with " + Columns.Count + " columns loaded");
                Console.WriteLine("1. Export to excel");
                Console.WriteLine("2. Change export settings");
                Console.WriteLine("3. Back");

                var option = Console.ReadLine();

                switch (option)
                {
                    case "1":
                        // If they do not have a config file then do not allow the export
                        if (!FileIo.IsEConfigExist)
                        {
                            Console.WriteLine("You do not have a configuration file in this directory, please set one up");
                            continue;
                        }
                        
                        // If they agree about the missing columns then continue
                        if (ValidateColumns())
                            return;
                        if (Columns.Count == 0)
                        {
                            Console.Clear();
                            Console.WriteLine("No Columns to export");
                            Console.WriteLine("Press enter to continue");
                            Console.ReadLine();
                            continue;
                        }
                        SaveToExcel();
                        break;

                    case "2":
                        break;

                    case "3":
                        return;

                    default:
                        continue;
                }
            }
        }

        /// <summary>
        /// Shows the columns that will not be exported to excel and asks the user for confirmation
        /// </summary>
        /// <returns>If the user would like to stop the export</returns>
        private static bool ValidateColumns()
        {
            Console.WriteLine("We have detected " + Columns.Count + " column properties");

            // Extract all of the apps that will not be exported
            var apps =
                Program.Applications[0].ProperyNames.Where(
                    t =>
                        Columns.Where(p => string.Equals(p.Name, t, StringComparison.CurrentCultureIgnoreCase))
                            .ToList()
                            .Count == 0).ToList();

            // If they do not have any apps to export then return false
            if (apps.Count == 0) return false;

            Console.WriteLine("Some columns are not going to be exported, this list is below");
            Console.WriteLine("If you would like to exit please type y");
            // Display all of the apps
            foreach (var app in apps)
            {
                Console.WriteLine(app);
            }

            // Return the result of there action
            return Console.ReadLine() == "y";
        }

        /// <summary>
        /// Saves all of the information stored to excel
        /// </summary>
        private static void SaveToExcel()
        {
            if (Columns.Count == 0)
                throw new Exception("No Columns to export");
            
            var formatAsTable = (Header.BackgroundColour != "") && (Header.TextColour != "");

            

            Console.WriteLine("Starting export to excel");
            var xlApp = new Application();

            object misValue = Missing.Value;

            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Worksheet) xlWorkBook.Worksheets.Item[1];

            var row = 0;
            for (var i = 0; i < Program.Applications.Count; i++)
            {
                if (Rows[row].Colour != "")
                {
                    formatAsTable = false;
                }
                row++;
                if (row >= Rows.Count)
                    row = 0;
            }

            if (formatAsTable)
            {
                Console.WriteLine("Adding format as table");
                var bottomRight = ExcelColumnFromNumber(Columns.Count);
                bottomRight = bottomRight + (Program.Applications.Count);
                var range = xlWorkSheet.Range["A1:" + bottomRight];
                xlWorkSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, range, Type.Missing,
                    Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo, Type.Missing).Name = "WFTableStyle";
                xlWorkSheet.ListObjects.Item["WFTableStyle"].TableStyle = "TableStyleMedium2";
            }

            Console.WriteLine("Inputting Column names");
            for (var i = 0; i < Columns.Count; i++)
            {
                xlWorkSheet.Cells[1, (i + 1)] = Columns[i].Name;
            }

            for (var i = 0; i < Program.Applications.Count; i++)
            {
                Console.WriteLine("Inputting item no " + (i + 1));
                for (var j = 0; j < Columns.Count; j++)
                {
                    for (var k = 0; k < Program.Applications[i].ProperyNames.Count; k++)
                    {
                        if (string.Equals(Program.Applications[i].ProperyNames[k], Columns[j].Name,
                            StringComparison.CurrentCultureIgnoreCase))
                        {
                            xlWorkSheet.Cells[(i + 2), (j + 1)] = Program.Applications[i].Properties[k];
                        }
                    }
                }
            }
            
            Console.WriteLine("Formatting all of the Columns");
            for (var i = 0; i < Columns.Count; i++)
            {
                ((Range) xlWorkSheet.Cells[1, (i + 1)]).EntireColumn.ColumnWidth = Columns[i].Width;
            }

            Console.WriteLine("Formatting the header");
            ((Range) xlWorkSheet.Cells[1, 1]).EntireRow.RowHeight = Header.Height;
            if (Header.BackgroundColour != "")
            {
                ((Range) xlWorkSheet.Cells[1, 1]).EntireRow.Interior.Color =
                    ColorTranslator.ToOle(ColorTranslator.FromHtml(Header.BackgroundColour));
                ((Range) xlWorkSheet.Cells[1, 1]).EntireRow.Font.Color =
                    ColorTranslator.ToOle(ColorTranslator.FromHtml(Header.TextColour));
            }

            Console.WriteLine("Formatting the rows");
            row = 0;
            for (var i = 0; i < Program.Applications.Count; i++)
            {
                ((Range) xlWorkSheet.Cells[(i + 2), 1]).EntireRow.RowHeight = Rows[row].Height;
                if (Rows[row].Colour != "")
                {
                    formatAsTable = false;
                    ((Range) xlWorkSheet.Cells[(i + 2), 1]).EntireRow.Interior.Color =
                        ColorTranslator.ToOle(ColorTranslator.FromHtml(Rows[row].Colour));
                }
                row++;
                if (row >= Rows.Count)
                    row = 0;
            }

            Console.WriteLine("Adding the OLE Objects");

            var oleObjects = (OLEObjects) xlWorkSheet.OLEObjects(Type.Missing);

            for (var i = 0; i < Program.Applications.Count; i++)
            {
                var application = Program.Applications[i];
                for (var index = 0; index < application.Properties.Count; index++)
                {
                    var property = application.Properties[index];
                    if (!property.Contains("||(|)||")) continue;
                    var split = property.Split(new[] {"||(|)||"}, StringSplitOptions.None);
                    for (var j = 1; j < split.Length; j = j + 2)
                    {
                        var filename = split[j].Split(new[] {"||()||"}, StringSplitOptions.None)[1];
                        decimal left = 0;
                        int column;
                        for (column = 0; column < Columns.Count; column++)
                        {
                            if (Columns[column].Name == application.ProperyNames[index])
                            {
                                break;
                            }
                            left += Columns[column].Width;
                        }

                        // If k is equal to the count then the attachment was not found, dont save it to excel
                        if (column == Columns.Count)
                        {
                            break;
                        }

                        var top = Header.Height;
                        for (var k = 0; k <= i; k++)
                        {
                            if (k == i)
                            {
                                top += Rows[(k%Rows.Count)].Height*(decimal) 0.1;
                            }
                            else
                            {
                                top += Rows[(k%Rows.Count)].Height;
                            }
                        }

                        var width = (double) (Columns[column].Width*5);
                        var height = (double) (Rows[i%Rows.Count].Height/(decimal) 1.2);
                        var ole = oleObjects.Add(Type.Missing,
                            Directory.GetCurrentDirectory() + "\\Objects\\" + filename, false, false, Type.Missing,
                            Type.Missing, Type.Missing, left*(decimal) 5.415, top, width/3, height);
                        ole.ShapeRange.LockAspectRatio = MsoTriState.msoFalse;
                        ole.Width = width;
                        ole.Height = height;
                    }
                }
            }

            //xlWorkSheet.Shapes.AddOLEObject(
            //    Directory.GetCurrentDirectory() +
            //    "\\Objects\\Apps AIS - list of apps gone through internal review 08022016.msg", 500, 500);

            xlWorkBook.SaveAs(Directory.GetCurrentDirectory() + "\\Export", XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, XlSaveAsAccessMode.xlNoChange,
                XlSaveConflictResolution.xlUserResolution,
                true, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
            Console.WriteLine("Excel save has finished");
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occured while releasing object " + ex);
            }
            finally
            {
                GC.Collect();
            }
        }

        /// <summary>
        /// http://stackoverflow.com/questions/837155/fastest-function-to-generate-excel-column-letters-in-c-sharp
        /// </summary>
        /// <param name="column">The integer excel column</param>
        /// <returns>Alphabet representation of the column</returns>
        private static string ExcelColumnFromNumber(int column)
        {
            var columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                var currentLetterNumber = (columnNumber - 1) % 26;
                var currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }
            return columnString;
        }
    }
}