/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Demo.Testing;
using NanoXLSX;
using Styles;

namespace Demo
{
    class Program
    {
        /// <summary>
        /// Method to run all demos / testing scenarios (currently disabled)
        /// </summary>
        /// <param name="args">Not used</param>
        static void Main(string[] args)
        {
            BasicDemo();
            Read();
            ShortenerDemo();
            StreamDemo();
            AsyncDemo(); // Normally, this method should be called with the await keyword (what is not possible here). Usually, async methods are called along the call stack with await until a terminal element (like a WPF button) is reached
            Demo1();
            Demo2();
            Demo3();
            Demo4();
            Demo5();
            Demo6();
            Demo7();
            Demo8();
            Demo9();
            Demo10();

            /* ### PERFORMANCE TESTS ### */
            // # Use tests in this section to test the performance of NanoXLSX
            /* ######################### */
            //   Performance.StressTest("stressTest.xlsx", "worksheet", 1500, 2);
            /* ######################### */
            /* ###### TYPE TESTS ####### */
            // # Use tests in this section to test the type casting of NanoXLSX
            /* ######################### */
            // Testing.TypeTesting.NumericTypeTesting("numericTest.xlsx");
            /* ######################### */
        }


        /// <summary>
        /// This is a very basic demo (adding three values and save the workbook)
        /// </summary>
        private static void BasicDemo()
        {
            Workbook workbook = new Workbook("basic.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");              // Add cell A1
            workbook.CurrentWorksheet.AddNextCell("Test2");             // Add cell B1
            workbook.CurrentWorksheet.AddNextCell("Test3");             // Add cell C1
            workbook.Save();
        }

        /// <summary>
        ///  This is a demo to read the previously created basix.xlsx file
        /// </summary>
        private static void Read()
        {
            Workbook wb = Workbook.Load("basic.xlsx");
            System.Console.WriteLine("contains worksheet name: " + wb.CurrentWorksheet.SheetName);
            foreach (KeyValuePair<string, Cell> cell in wb.CurrentWorksheet.Cells)
            {
                System.Console.WriteLine("Cell address: " + cell.Key + ": content:'" + cell.Value.Value + "'");
            }

            // The same as stream
            Workbook wb2 = null;
            using (FileStream fs = new FileStream("basic.xlsx", FileMode.Open))
            {
                wb2 = Workbook.Load(fs);
            }
            System.Console.WriteLine("contains worksheet name: " + wb2.CurrentWorksheet.SheetName);
            foreach (KeyValuePair<string, Cell> cell in wb2.CurrentWorksheet.Cells)
            {
                System.Console.WriteLine("Cell address: " + cell.Key + ": content:'" + cell.Value.Value + "'");
            }
        }


        /// <summary>
        /// This method shows the shortened style of writing cells
        /// </summary>
        private static void ShortenerDemo()
        {
            Workbook wb = new Workbook("shortenerDemo.xlsx", "Sheet1"); // Create a workbook (important: A worksheet must be created as well) 
            wb.WS.Value("Some Text");                                   // Add cell A1
            wb.WS.Value(58.55, BasicStyles.DoubleUnderline);            // Add a formatted value to cell B1
            wb.WS.Right(2);                                             // Move to cell E1   
            wb.WS.Value(true);                                          // Add cell E1
            wb.AddWorksheet("Sheet2");                                  // Add a new worksheet
            wb.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;    // Change the cell direction
            wb.WS.Value("This is another text");                        // Add cell A1
            wb.WS.Formula("=A1");                                       // Add a formula in Cell A2
            wb.WS.Down();                                               // Go to cell A4
            wb.WS.Value("Formatted Text", BasicStyles.Bold);            // Add a formatted value to cell A4
            wb.Save();                                                  // Save the workbook
        }

        /// <summary>
        /// This method shows how to save a workbook as stream 
        /// </summary>
        private static void StreamDemo()
        {
            Workbook workbook = new Workbook(true);                         // Create new workbook without file name
            workbook.CurrentWorksheet.AddNextCell("This is an example");    // Add cell A1
            workbook.CurrentWorksheet.AddNextCellFormula("=A1");            // Add formula in cell B1
            workbook.CurrentWorksheet.AddNextCell(123456789);               // Add cell C1
            FileStream fs = new FileStream("stream.xlsx", FileMode.Create); // Create a file stream (could also be a memory stream or whatever writable stream you want)
            workbook.SaveAsStream(fs);                                      // Save the workbook into the stream
        }


        /// <summary>
        /// This method shows how to save a workbook asynchronous
        /// </summary>
        private static async Task AsyncDemo()
        {
            Workbook workbook = new Workbook("async.xlsx", "shet1");       // Create new workbook with file name
            workbook.WS.Value("Some text");                                 // Add cell A1
            workbook.WS.Value(222);                                         // Add cell B1
            workbook.WS.Formula("=A2");                                     // Add cell C1
            await workbook.SaveAsync();                                     // Save async
        }

        /// <summary>
        /// This method shows the usage of AddNextCell with several data types and formulas. Furthermore, the several types of Addresses are demonstrated
        /// </summary>
        private static void Demo1()
        {
            Workbook workbook = new Workbook("test1.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");              // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(123);                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(true);                // Add cell C1
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(123.456d);            // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(123.789f);            // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);        // Add cell C2
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 3
            workbook.CurrentWorksheet.AddNextCellFormula("B1*22");      // Add cell A3 as formula (B1 times 22)
            workbook.CurrentWorksheet.AddNextCellFormula("ROUNDDOWN(A2,1)"); // Add cell B3 as formula (Floor A2 with one decimal place)
            workbook.CurrentWorksheet.AddNextCellFormula("PI()");       // Add cell C3 as formula (Pi = 3.14.... )
            workbook.AddWorksheet("Addresses");                                                 // Add new worksheet
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.Disabled;  // Disable automatic addressing
            workbook.CurrentWorksheet.AddCell("Default", 0, 0);                                 // Add a value
            Address address = new Address(1, 0, Cell.AddressType.Default);                      // Create Address with default behavior
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 0);                        // Add the string of the address
            workbook.CurrentWorksheet.AddCell("Fixed Column", 0, 1);                            // Add a value
            address = new Address(1, 1, Cell.AddressType.FixedColumn);                          // Create Address with fixed column
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 1);                        // Add the string of the address
            workbook.CurrentWorksheet.AddCell("Fixed Row", 0, 2);                               // Add a value
            address = new Address(1, 2, Cell.AddressType.FixedRow);                             // Create Address with fixed row
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 2);                        // Add the string of the address
            workbook.CurrentWorksheet.AddCell("Fixed Row and Column", 0, 3);                    // Add a value
            address = new Address(1, 3, Cell.AddressType.FixedRowAndColumn);                    // Create Address with fixed row and column
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 3);                        // Add the string of the address
            workbook.Save();                                                                    // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of several data types, the method AddCell, more than one worksheet and the SaveAs method
        /// </summary>
        private static void Demo2()
        {
            Workbook workbook = new Workbook(false);                    // Create new workbook
            workbook.AddWorksheet("Sheet1");                            // Add a new Worksheet and set it as current sheet
            workbook.CurrentWorksheet.AddNextCell("月曜日");            // Add cell A1 (Unicode)
            workbook.CurrentWorksheet.AddNextCell(-987);                // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(false);               // Add cell C1
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(-123.456d);           // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(-123.789f);           // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);        // Add cell C3
            workbook.AddWorksheet("Sheet2");                            // Add a new Worksheet and set it as current sheet
            workbook.CurrentWorksheet.AddCell("ABC", "A1");             // Add cell A1
            workbook.CurrentWorksheet.AddCell(779, 2, 1);               // Add cell C2 (zero based addresses: column 2=C, row 1=2)
            workbook.CurrentWorksheet.AddCell(false, 3, 2);             // Add cell D3 (zero based addresses: column 3=D, row 2=3)
            workbook.CurrentWorksheet.AddNextCell(0);                   // Add cell E3 (direction: column to column)
            List<object> values = new List<object>() { "V1", true, 16.8 }; // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, "A4:C4");    // Add a cell range to A4 - C4
            workbook.SaveAs("test2.xlsx");                              // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of flipped direction when using AddNextCell, reading of the current cell address, and retrieving of cell values
        /// </summary>
        private static void Demo3()
        {
            Workbook workbook = new Workbook("test3.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;  // Change the cell direction
            workbook.CurrentWorksheet.AddNextCell(1);                   // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(2);                   // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(3);                   // Add cell A3
            workbook.CurrentWorksheet.AddNextCell(4);                   // Add cell A4
            int row = workbook.CurrentWorksheet.GetCurrentRowNumber(); // Get the row number (will be 4 = row 5)
            int col = workbook.CurrentWorksheet.GetCurrentColumnNumber(); // Get the column number (will be 0 = column A)
            workbook.CurrentWorksheet.AddNextCell("This cell has the row number " + (row + 1) + " and column number " + (col + 1));
            workbook.CurrentWorksheet.GoToNextColumn();                 // Go to Column B
            workbook.CurrentWorksheet.AddNextCell("A");                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell("B");                 // Add cell B2
            workbook.CurrentWorksheet.AddNextCell("C");                 // Add cell B3
            workbook.CurrentWorksheet.AddNextCell("D");                 // Add cell B4
            workbook.CurrentWorksheet.RemoveCell("A2");                 // Delete cell A2
            workbook.CurrentWorksheet.RemoveCell(1, 1);                  // Delete cell B2
            workbook.CurrentWorksheet.GoToNextRow(3);                   // Move 3 rows down
            object value = workbook.CurrentWorksheet.GetCell(1, 2).Value;  // Gets the value of cell B3
            workbook.CurrentWorksheet.AddNextCell("Value of B3 is: " + value);
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.Disabled;   // Disable automatic cell addressing
            workbook.CurrentWorksheet.AddCell("Text A", 3, 0);          // Add manually placed value
            workbook.CurrentWorksheet.AddCell("Text B", 4, 1);          // Add manually placed value
            workbook.CurrentWorksheet.AddCell("Text C", 3, 2);          // Add manually placed value
            workbook.Save();                                            // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of several styles, column widths and row heights
        /// </summary>
        private static void Demo4()
        {
            Workbook workbook = new Workbook("test4.xlsx", "Sheet1");                                        // Create new workbook
            List<object> values = new List<object>() { "Header1", "Header2", "Header3" };                    // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, new Address(0, 0), new Address(2, 0));    // Add a cell range to A4 - C4
            workbook.CurrentWorksheet.Cells["A1"].SetStyle(BasicStyles.Bold);                          // Assign predefined basic style to cell
            workbook.CurrentWorksheet.Cells["B1"].SetStyle(BasicStyles.Bold);                          // Assign predefined basic style to cell
            workbook.CurrentWorksheet.Cells["C1"].SetStyle(BasicStyles.Bold);                          // Assign predefined basic style to cell
            workbook.CurrentWorksheet.GoToNextRow();                                                         // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);                                             // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(2);                                                        // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(3);                                                        // Add cell B2
            workbook.CurrentWorksheet.GoToNextRow();                                                         // Go to Row 3
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now.AddDays(1));                                  // Add cell B1
            workbook.CurrentWorksheet.AddNextCell("B");                                                      // Add cell B2
            workbook.CurrentWorksheet.AddNextCell("C");                                                      // Add cell B3

            Style s = new Style();                                                                          // Create new style
            s.CurrentFill.SetColor("FF22FF11", Fill.FillType.fillColor);                              // Set fill color
            s.CurrentFont.DoubleUnderline = true;                                                           // Set double underline
            s.CurrentCellXf.HorizontalAlign = CellXf.HorizontalAlignValue.center;                     // Set alignment

            Style s2 = s.CopyStyle();                                                                       // Copy the previously defined style
            s2.CurrentFont.Italic = true;                                                                   // Change an attribute of the copied style

            workbook.CurrentWorksheet.Cells["B2"].SetStyle(s);                                              // Assign style to cell
            workbook.CurrentWorksheet.GoToNextRow();                                                        // Go to Row 3
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now.AddDays(2));                                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(true);                                                    // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(false, s2);                                               // Add cell B3 with style in the same step 
            workbook.CurrentWorksheet.Cells["C2"].SetStyle(BasicStyles.BorderFrame);                  // Assign predefined basic style to cell

            Style s3 = BasicStyles.Strike;                                                            // Create a style from a predefined style
            s3.CurrentCellXf.TextRotation = 45;                                                             // Set text rotation
            s3.CurrentCellXf.VerticalAlign = CellXf.VerticalAlignValue.center;                        // Set alignment

            workbook.CurrentWorksheet.Cells["B4"].SetStyle(s3);                                             // Assign style to cell

            workbook.CurrentWorksheet.SetColumnWidth(0, 20f);                                               // Set column width
            workbook.CurrentWorksheet.SetColumnWidth(1, 15f);                                               // Set column width
            workbook.CurrentWorksheet.SetColumnWidth(2, 25f);                                               // Set column width
            workbook.CurrentWorksheet.SetRowHeight(0, 20);                                                 // Set row height
            workbook.CurrentWorksheet.SetRowHeight(1, 30);                                                 // Set row height

            workbook.Save();                                                                               // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of cell ranges, adding and removing styles, and meta data 
        /// </summary>
        private static void Demo5()
        {
            Workbook workbook = new Workbook("test5.xlsx", "Sheet1");                                   // Create new workbook
            List<object> values = new List<object>() { "Header1", "Header2", "Header3" };               // Create a List of values
            workbook.CurrentWorksheet.SetActiveStyle(BasicStyles.BorderFrameHeader);              // Assign predefined basic style as active style
            workbook.CurrentWorksheet.AddCellRange(values, "A1:C1");                                    // Add cell range

            values = new List<object>() { "Cell A2", "Cell B2", "Cell C2" };                            // Create a List of values
            workbook.CurrentWorksheet.SetActiveStyle(BasicStyles.BorderFrame);                    // Assign predefined basic style as active style
            workbook.CurrentWorksheet.AddCellRange(values, "A2:C2");                                    // Add cell range (using active style)

            values = new List<object>() { "Cell A3", "Cell B3", "Cell C3" };                            // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, "A3:C3");                                    // Add cell range (using active style)

            values = new List<object>() { "Cell A4", "Cell B4", "Cell C4" };                            // Create a List of values
            workbook.CurrentWorksheet.ClearActiveStyle();                                               // Clear the active style 
            workbook.CurrentWorksheet.AddCellRange(values, "A4:C4");                                    // Add cell range (without style)

            workbook.WorkbookMetadata.Title = "Test 5";                                                 // Add meta data to workbook
            workbook.WorkbookMetadata.Subject = "This is the 5th NanoXLSX test";                        // Add meta data to workbook
            workbook.WorkbookMetadata.Creator = "NanoXLSX";                                             // Add meta data to workbook
            workbook.WorkbookMetadata.Keywords = "Keyword1;Keyword2;Keyword3";                          // Add meta data to workbook

            workbook.Save();                                                                            // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of merging cells, protecting cells, worksheet password protection and workbook protection
        /// </summary>
        private static void Demo6()
        {
            Workbook workbook = new Workbook("test6.xlsx", "Sheet1");                                   // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Merged1");                                           // Add cell A1
            workbook.CurrentWorksheet.MergeCells("A1:C1");                                              // Merge cells from A1 to C1
            workbook.CurrentWorksheet.GoToNextRow();                                                    // Go to next row
            workbook.CurrentWorksheet.AddNextCell(false);                                               // Add cell A2
            workbook.CurrentWorksheet.MergeCells("A2:D2");                                              // Merge cells from A2 to D1
            workbook.CurrentWorksheet.GoToNextRow();                                                    // Go to next row
            workbook.CurrentWorksheet.AddNextCell("22.2d");                                             // Add cell A3
            workbook.CurrentWorksheet.MergeCells("A3:E4");                                              // Merge cells from A3 to E4
            workbook.AddWorksheet("Protected");                                                         // Add a new worksheet
            workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.sort);               // Allow to sort sheet (worksheet is automatically set as protected)
            workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.insertRows);         // Allow to insert rows
            workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.selectLockedCells);  // Allow to select cells (locked cells caused automatically to select unlocked cells)
            workbook.CurrentWorksheet.AddNextCell("Cell A1");                                           // Add cell A1
            workbook.CurrentWorksheet.AddNextCell("Cell B1");                                           // Add cell B1
            workbook.CurrentWorksheet.Cells["A1"].SetCellLockedState(false, true);                      // Set the locking state of cell A1 (not locked but value is hidden when cell selected)
            workbook.AddWorksheet("PWD-Protected");                                                     // Add a new worksheet
            workbook.CurrentWorksheet.AddCell("This worksheet is password protected. The password is:", 0, 0);  // Add cell A1
            workbook.CurrentWorksheet.AddCell("test123", 0, 1);                                         // Add cell A2
            workbook.CurrentWorksheet.SetSheetProtectionPassword("test123");                            // Set the password "test123"
            workbook.SetWorkbookProtection(true, true, true, null);                                     // Set workbook protection (windows locked, structure locked, no password)
            workbook.Save();                                                                            // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of hiding rows and columns, auto-filter and worksheet name sanitizing
        /// </summary>
        private static void Demo7()
        {
            Workbook workbook = new Workbook(false);                                                    // Create new workbook without worksheet
            string invalidSheetName = "Sheet?1";                                                        // ? is not allowed in the names of worksheets
            string sanitizedSheetName = Worksheet.SanitizeWorksheetName(invalidSheetName, workbook);    // Method to sanitize a worksheet name (replaces ? with _)
            workbook.AddWorksheet(sanitizedSheetName);                                                  // Add new worksheet
            Worksheet ws = workbook.CurrentWorksheet;                                                   // Create reference (shortening)
            List<object> values = new List<object>() { "Cell A1", "Cell B1", "Cell C1", "Cell D1" };    // Create a List of values
            ws.AddCellRange(values, "A1:D1");                                                           // Insert cell range
            values = new List<object>() { "Cell A2", "Cell B2", "Cell C2", "Cell D2" };                 // Create a List of values
            ws.AddCellRange(values, "A2:D2");                                                           // Insert cell range
            values = new List<object>() { "Cell A3", "Cell B3", "Cell C3", "Cell D3" };                 // Create a List of values
            ws.AddCellRange(values, "A3:D3");                                                           // Insert cell range
            ws.AddHiddenColumn("C");                                                                    // Hide column C
            ws.AddHiddenRow(1);                                                                         // Hider row 2 (zero-based: 1)
            ws.SetAutoFilter(1, 3);                                                                     // Set auto-filter for column B to D
            workbook.SaveAs("test7.xlsx");                                                              // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of cell and worksheet selection, auto-sanitizing of worksheet names
        /// </summary>
        private static void Demo8()
        {
            Workbook workbook = new Workbook("test8.xlsx", "Sheet*1", true);  				            // Create new workbook with invalid sheet name (*); Auto-Sanitizing will replace * with _
            workbook.CurrentWorksheet.AddNextCell("Test");              								// Add cell A1
            workbook.CurrentWorksheet.SetSelectedCells("A5:B10");										// Set the selection to the range A5:B10
            workbook.AddWorksheet("Sheet2");															// Create new worksheet
            workbook.CurrentWorksheet.AddNextCell("Test2");              								// Add cell A1
            Range range = new Range(new Address(1, 1), new Address(3, 3));			// Create a cell range for the selection B2:D4
            workbook.CurrentWorksheet.SetSelectedCells(range);											// Set the selection to the range
            workbook.AddWorksheet("Sheet2", true);							// Create new worksheet with already existing name; The name will be changed to Sheet21 due to auto-sanitizing (appending of 1)
            workbook.CurrentWorksheet.AddNextCell("Test3");              								// Add cell A1
            workbook.CurrentWorksheet.SetSelectedCells(new Address(2, 2), new Address(4, 4));	// Set the selection to the range C3:E5
            workbook.SetSelectedWorksheet(1);															// Set the second Tab as selected (zero-based: 1)
            workbook.Save();                                            								// Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of basic Excel formulas
        /// </summary>
        private static void Demo9()
        {
            Workbook workbook = new Workbook("test9.xlsx", "sheet1");                                   // Create a new workbook 
            List<object> numbers = new List<object> { 1.15d, 2.225d, 13.8d, 15d, 15.1d, 17.22d, 22d, 107.5d, 128d }; // Create a list of numbers
            List<object> texts = new List<object>() { "value 1", "value 2", "value 3", "value 4", "value 5", "value 6", "value 7", "value 8", "value 9" }; // Create a list of strings (for vlookup)
            workbook.WS.Value("Numbers", BasicStyles.Bold);                                       // Add a header with a basic style
            workbook.WS.Value("Values", BasicStyles.Bold);                                        // Add a header with a basic style
            workbook.WS.Value("Formula type", BasicStyles.Bold);                                  // Add a header with a basic style
            workbook.WS.Value("Formula value", BasicStyles.Bold);                                 // Add a header with a basic style
            workbook.WS.Value("(See also worksheet2)");                                                 // Add a note
            workbook.CurrentWorksheet.AddCellRange(numbers, "A2:A10");                                  // Add the numbers as range
            workbook.CurrentWorksheet.AddCellRange(texts, "B2:B10");                                    // Add the values as range

            workbook.CurrentWorksheet.SetCurrentCellAddress("D2");                                      // Set the "cursor" to D2
            Cell c;                                                                                     // Create an empty cell object (reusable)
            c = BasicFormulas.Average(new Range("A2:A10"));                                   // Define an average formula
            workbook.CurrentWorksheet.AddCell("Average", "C2");                                         // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D2");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Ceil(new Address("A2"), 0);                                     // Define a ceil formula
            workbook.CurrentWorksheet.AddCell("Ceil", "C3");                                           // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D3");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Floor(new Address("A2"), 0);                                    // Define a floor formula
            workbook.CurrentWorksheet.AddCell("Floor", "C4");                                           // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D4");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Round(new Address("A3"), 1);                                    // Define a round formula with one digit after the comma
            workbook.CurrentWorksheet.AddCell("Round", "C5");                                           // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D5");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Max(new Range("A2:A10"));                                       // Define a max formula
            workbook.CurrentWorksheet.AddCell("Max", "C6");                                             // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D6");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Min(new Range("A2:A10"));                                       // Define a min formula
            workbook.CurrentWorksheet.AddCell("Min", "C7");                                             // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D7");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Median(new Range("A2:A10"));                                    // Define a median formula
            workbook.CurrentWorksheet.AddCell("Median", "C8");                                          // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D8");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Sum(new Range("A2:A10"));                                       // Define a sum formula
            workbook.CurrentWorksheet.AddCell("Sum", "C9");                                             // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D9");                                                 // Add the formula to the worksheet

            c = BasicFormulas.VLookup(13.8d, new Range("A2:B10"), 2, true);                   // Define a vlookup formula (look for the value of the number 13.8) 
            workbook.CurrentWorksheet.AddCell("Vlookup", "C10");                                        // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D10");                                                // Add the formula to the worksheet

            workbook.AddWorksheet("sheet2");                                                            // Create a new worksheet
            c = BasicFormulas.VLookup(workbook.Worksheets[0], new Address("B4"), workbook.Worksheets[0], new Range("B2:C10"), 2, true); // Define a vlookup formula in worksheet1 (look for the text right of the (value of) cell B4) 
            workbook.WS.Value(c);                                                                       // Add the formula to the worksheet

            c = BasicFormulas.Median(workbook.Worksheets[0], new Range("A2:A10"));            // Define a median formula in worksheet1
            workbook.WS.Value(c);                                                                       // Add the formula to the worksheet

            workbook.Save();                                                                            // Save the workbook
        }

        /// <summary>
        /// This demo shows the usage of style appending
        /// </summary>
        private static void Demo10()
        {
            Workbook wb = new Workbook("demo10.xlsx", "styleAppending");                                // Create a new workbook

            Style style = new Style();                                                                  // Create a new style
            style.Append(BasicStyles.Bold);                                                       // Append a basic style (bold) 
            style.Append(BasicStyles.Underline);                                                  // Append a basic style (underline) 
            style.Append(BasicStyles.Font("Arial Black", 20));                                    // Append a basic style (custom font) 

            wb.WS.Value("THIS IS A TEST", style);                                                       // Add text and the appended style
            wb.WS.Down();                                                                               // Go to a new row

            Style chainedStyle = new Style()                                                            // Create a new style...
                .Append(BasicStyles.Underline)                                                    // ... and append another part (chaining underline)
                .Append(BasicStyles.ColorizedText("FF00FF"))                                      // ... and append another part (chaining colorized text)
                .Append(BasicStyles.ColorizedBackground("AAFFAA"));                               // ... and append another part (chaining colorized background)

            wb.WS.Value("Another test", chainedStyle);                                                  // Add text and the appended style

            wb.Save();                                                                                  // Save the workbook
        }

    }
}
