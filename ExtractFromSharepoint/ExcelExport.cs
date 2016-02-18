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
    internal static class ExcelExport
    {
        internal static List<ExcelColumn> Columns = new List<ExcelColumn>();

        internal static Header Header = new Header();

        internal static List<Row> Rows = new List<Row>();

        internal static void Main()
        {
            FileIo.ImportExcelConfig();
            if (ValidateColumns())
                return;
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

        private static bool ValidateColumns()
        {
            Console.WriteLine("We have detected " + Columns.Count + " column properties");
            var apps = new List<string>();
            foreach (string t in Program.Applications[0].ProperyNames)
            {
                if (Columns.Where(p => p.Name == t).ToList().Count == 0)
                {
                    apps.Add(t);
                }
            }

            if (apps.Count > 0)
            {
                Console.WriteLine("Some columns are not going to be exported, this list is below");
                Console.WriteLine("If you would like to exit please type y");
                foreach (var app in apps)
                {
                    Console.WriteLine(app);
                }
                if (Console.ReadLine() == "y")
                {
                    return true;
                }
            }
            return false;
        }

        static void SaveToExcel()
        {
            decimal dpiX;
            decimal dpiY;
            using (var graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                dpiX = (decimal) graphics.DpiX;
                dpiY = (decimal) graphics.DpiY;
            }

            Console.WriteLine("Starting export to excel");
            var xlApp = new Application();

            object misValue = Missing.Value;

            var xlWorkBook = xlApp.Workbooks.Add(misValue);
            var xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.Item[1];

            Console.WriteLine("Inputting Column names");
            for (var i = 0; i < Columns.Count; i++)
            {
                xlWorkSheet.Cells[1, (i + 1)] = Columns[i].Name;
            }

            for (var i = 0; i < Program.Applications.Count; i++)
            {
                Console.WriteLine("Inputting item no " + (i+1));
                for (var j = 0; j < Columns.Count; j++)
                {
                    for (var k = 0; k < Program.Applications[i].ProperyNames.Count; k++)
                    {
                        if (Program.Applications[i].ProperyNames[k] == Columns[j].Name)
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
            ((Range)xlWorkSheet.Cells[1, 1]).EntireRow.RowHeight = Header.Height;
            ((Range)xlWorkSheet.Cells[1, 1]).EntireRow.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(Header.BackgroundColour));
            ((Range)xlWorkSheet.Cells[1, 1]).EntireRow.Font.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(Header.TextColour));

            Console.WriteLine("Formatting the rows");
            var row = 0;
            for (var i = 0; i < Program.Applications.Count; i++)
            {
                ((Range)xlWorkSheet.Cells[(i + 2), 1]).EntireRow.RowHeight = Rows[row].Height;
                ((Range)xlWorkSheet.Cells[(i + 2), 1]).EntireRow.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml(Rows[row].Colour));
                row++;
                if (row >= Rows.Count)
                    row = 0;
            }

            Console.WriteLine("Adding the OLE Objects");

            var oleObjects = (OLEObjects) xlWorkSheet.OLEObjects(Type.Missing);

            for (int i = 0; i < Program.Applications.Count; i++)
            {
                var application = Program.Applications[i];
                for (int index = 0; index < application.Properties.Count; index++)
                {
                    var property = application.Properties[index];
                    if (property.Contains("||(|)||"))
                    {
                        var split = property.Split(new string[] {"||(|)||"}, StringSplitOptions.None);
                        for (int j = 1; j < split.Length; j=j+2)
                        {
                            var filename = split[j].Split(new string[] {"||()||"}, StringSplitOptions.None)[1];
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
                                    top += Rows[(k % Rows.Count)].Height*(decimal) 0.1;
                                }
                                else
                                {
                                    top += Rows[(k % Rows.Count)].Height;
                                }
                            }

                            var width = (double) (Columns[column].Width*(decimal) 5);
                            var height = (double) (Rows[i%Rows.Count].Height/(decimal) 1.2);
                            var ole = oleObjects.Add(Type.Missing, Directory.GetCurrentDirectory() + "\\Objects\\" + filename, false, false, Type.Missing,
                                Type.Missing, Type.Missing, left*(decimal) 5.415, top, width/3, height);
                            ole.ShapeRange.LockAspectRatio = MsoTriState.msoFalse;
                            ole.Width = width;
                            ole.Height = height;
                        }
                    }
                }
            }

            //xlWorkSheet.Shapes.AddOLEObject(
            //    Directory.GetCurrentDirectory() +
            //    "\\Objects\\Apps AIS - list of apps gone through internal review 08022016.msg", 500, 500);

            xlWorkBook.SaveAs(Directory.GetCurrentDirectory() + "\\Export", XlFileFormat.xlOpenXMLWorkbook, misValue,
                misValue, misValue, misValue, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlUserResolution,
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
