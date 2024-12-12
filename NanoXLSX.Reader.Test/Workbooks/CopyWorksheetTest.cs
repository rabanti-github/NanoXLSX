using NanoXLSX;
using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX_Test.Workbooks
{
    public class CopyWorksheetTest
    {

        [Fact(DisplayName = "Test of the 'CopyWorksheetTo' function for proper saving")]
        public void CopyWorksheetSaveTest()
        {
            Workbook workbook1 = new Workbook("worksheet1");
            Workbook workbook2 = new Workbook("worksheet1b");
            Worksheet worksheet2 = createWorksheet();
            worksheet2.SheetName = "worksheet2";
            workbook1.AddWorksheet(worksheet2);
            Workbook.CopyWorksheetTo(worksheet2, "copy", workbook2);

            Workbook newWorkbook = TestUtils.WriteAndReadWorkbook(workbook2);
            Assert.Equal(workbook2.Worksheets.Count, newWorkbook.Worksheets.Count);
        }

        private Worksheet createWorksheet()
        {
            Worksheet w = new Worksheet();
            Style s1 = BasicStyles.BoldItalic;
            Style s2 = BasicStyles.Bold.Append(BasicStyles.DateFormat);
            w.AddCell("A1", "A1", s1);
            w.AddCell(true, "B2");
            w.AddCell(100, "C3", s2);
            w.AddCell(2.23f, "D4");
            w.AddCell(false, "D5");
            w.AddCellFormula("=A2", "E5");
            w.SetColumnWidth(2, 31.2f);
            w.SetRowHeight(2, 50.6f);
            w.AddHiddenColumn(1);
            w.AddHiddenColumn(3);
            w.SetColumnDefaultStyle(1, BasicStyles.Font("Comic Sans", 42));
            w.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.sort);
            w.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.autoFilter);
            w.SetSheetProtectionPassword("pwd");
            w.AddHiddenRow(1);
            w.AddHiddenRow(3);
            w.CurrentCellDirection = Worksheet.CellDirection.Disabled;
            w.DefaultColumnWidth = 55.5f;
            w.DefaultRowHeight = 45.3f;
            w.Hidden = true;
            w.MergeCells(new NanoXLSX.Range("D4:D5"));
            w.SetActiveStyle(s2);
            w.SetAutoFilter("B1:C2");
            w.SetCurrentCellAddress("D5");
            w.AddSelectedCells(new NanoXLSX.Range("C3:C3"));
            w.UseSheetProtection = true;
            w.SetSplit(3, 2, true, new Address("F4"), Worksheet.WorksheetPane.bottomRight);
            return w;
        }
    }
}
