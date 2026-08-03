using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using NanoXLSX.Colors;
using NanoXLSX.Exceptions;
using NanoXLSX.Internal.Readers;
using NanoXLSX.Internal.Structures;
using NanoXLSX.Internal.Writers;
using NanoXLSX.Registry;
using NanoXLSX.Styles;
using NanoXLSX.Test.Writer_Reader.Utils;
using Xunit;

namespace NanoXLSX.Test.Writer_Reader.WorkbookTest
{
    public class WorkbookWriteReadTest
    {
        [Fact(DisplayName = "Test of the correct processing of 3 strings when saving and loading a workbook")]
        public void StringProcessingTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell("Text1", "A1");
            workbook.CurrentWorksheet.AddCell("Text2", "A2");
            workbook.CurrentWorksheet.AddCell("", "A3");
            workbook.CurrentWorksheet.AddCell(null, "A4");
            workbook.CurrentWorksheet.AddCell("Text1", "A5");
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A1"].DataType);
            Assert.Equal("Text1", givenWorkbook.CurrentWorksheet.Cells["A1"].Value.ToString());
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A2"].DataType);
            Assert.Equal("Text2", givenWorkbook.CurrentWorksheet.Cells["A2"].Value.ToString());
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A3"].DataType);
            Assert.Equal("", givenWorkbook.CurrentWorksheet.Cells["A3"].Value.ToString());
            Assert.Equal(Cell.CellType.Empty, givenWorkbook.CurrentWorksheet.Cells["A4"].DataType);
            Assert.Null(givenWorkbook.CurrentWorksheet.Cells["A4"].Value);
            Assert.Equal(Cell.CellType.String, givenWorkbook.CurrentWorksheet.Cells["A5"].DataType);
            Assert.Equal("Text1", givenWorkbook.CurrentWorksheet.Cells["A5"].Value.ToString());
        }

        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property on a ARGB value, when writing and reading a workbook")]
        public void ReadMruColorsTest()
        {
            Workbook workbook = new Workbook();
            string color1 = "AACC00";
            string color2 = "FFDD22";
            workbook.AddMruColor(color1);
            workbook.AddMruColor(color2);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Equal(2, mruColors.Count);
            Assert.Equal("FF" + color1, mruColors[0].GetArgbValue());
            Assert.Equal("FF" + color2, mruColors[1].GetArgbValue());
        }


        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property on a indexed color, when writing and reading a workbook")]
        public void ReadMruColorsTest2()
        {
            Workbook workbook = new Workbook();
            workbook.AddMruColor(IndexedColor.Value.Blue4);
            workbook.AddMruColor(IndexedColor.Value.StrongYellow);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Equal(2, mruColors.Count);
            Assert.Equal(IndexedColor.GetArgbValue(IndexedColor.Value.Blue4), mruColors[0].GetArgbValue());
            Assert.Equal(IndexedColor.GetArgbValue(IndexedColor.Value.StrongYellow), mruColors[1].GetArgbValue());
        }


        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property when writing and reading a workbook, neglecting the default color")]
        public void ReadMruColorsTest3()
        {
            Workbook workbook = new Workbook();
            string color1 = "AACC00";
            string color2 = Fill.DefaultColor.RgbColor.ColorValue; // Should not be added (black / default color)
            workbook.AddMruColor(color1);
            workbook.AddMruColor(color2);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            mruColors.Sort();
            Assert.Single(mruColors);
            Assert.Equal("FF" + color1, mruColors[0].GetArgbValue());
        }

        [Fact(DisplayName = "Test of the (virtual) 'MruColors' property when writing and reading a workbook, neglecting an undefined color")]
        public void ReadMruColorsTest4()
        {
            Workbook workbook = new Workbook();
            Color color = Color.CreateNone();
            workbook.AddMruColor(color);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            List<Color> mruColors = ((List<Color>)givenWorkbook.GetMruColors());
            Assert.Empty(mruColors);
        }

        [Theory(DisplayName = "Test of the 'Hidden' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookHiddenTest(bool hidden)
        {
            Workbook workbook = new Workbook
            {
                Hidden = hidden
            };
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(hidden, givenWorkbook.Hidden);
        }

        [Theory(DisplayName = "Test of the 'SelectedWorksheet' property when writing and reading a workbook")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(2)]
        public void ReadWorkbookSelectedWorksheetTest(int index)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddWorksheet("sheet2");
            workbook.AddWorksheet("sheet3");
            workbook.AddWorksheet("sheet4");
            workbook.SetSelectedWorksheet(index);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(index, givenWorkbook.SelectedWorksheet);
        }

        [Theory(DisplayName = "Test of the 'LockWindowsIfProtected' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookLockWindowsTest(bool locked)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.SetWorkbookProtection(true, locked, false, null);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(locked, givenWorkbook.LockWindowsIfProtected);
        }

        [Theory(DisplayName = "Test of the 'LockStructureIfProtected' property when writing and reading a workbook")]
        [InlineData(true)]
        [InlineData(false)]
        public void ReadWorkbookLockStructureTest(bool locked)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.SetWorkbookProtection(true, false, locked, null);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(locked, givenWorkbook.LockStructureIfProtected);
        }

        [Fact(DisplayName = "Test that a workbook without defined names produces no defined names after round-trip")]
        public void DefinedNames_EmptyRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(1, "A1");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Empty(given.GetDefinedNames());
        }

        [Fact(DisplayName = "Test of a workbook-scoped defined name with a single cell reference")]
        public void DefinedNames_WorkbookScope_CellReference()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(42, "A1");
            workbook.AddDefinedNameCell("MyCell", workbook.CurrentWorksheet, "A1");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Single(given.GetDefinedNames());
            DefinedName dn = given.GetDefinedName("MyCell");
            Assert.Equal(DefinedName.NameType.Cell, dn.Type);
            Assert.Equal("$A$1", dn.TextValue);
            Assert.Same(given.CurrentWorksheet, dn.TargetWorksheet);
            Assert.Null(dn.LocalSheet);
            Assert.Null(dn.Comment);
        }

        [Fact(DisplayName = "Test of a workbook-scoped defined name with a range reference")]
        public void DefinedNames_WorkbookScope_RangeReference()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedNameRange("MyRange", workbook.CurrentWorksheet, "A1:B3");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            DefinedName dn = given.GetDefinedName("MyRange");
            Assert.Equal(DefinedName.NameType.Range, dn.Type);
            Assert.Equal("$A$1:$B$3", dn.TextValue);
            Assert.Same(given.CurrentWorksheet, dn.TargetWorksheet);
        }

        [Fact(DisplayName = "Test of a workbook-scoped defined name holding a formula expression")]
        public void DefinedNames_WorkbookScope_FormulaReference()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedNameFormula("MySum", "SUM(sheet1!$A$1:$A$3)");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            DefinedName dn = given.GetDefinedName("MySum");
            Assert.NotNull(dn);
            Assert.Equal("SUM(sheet1!$A$1:$A$3)", dn.TextValue);
        }

        [Fact(DisplayName = "Test of a worksheet-scoped defined name (localSheetId) round-trip")]
        public void DefinedNames_WorksheetScope()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddWorksheet("sheet2");
            workbook.AddDefinedNameCell("LocalName", workbook.Worksheets[1], "B2", workbook.Worksheets[1]);
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Single(given.GetDefinedNames());
            Worksheet sheet2 = given.GetWorksheet("sheet2");
            DefinedName dn = given.GetDefinedName("LocalName", sheet2);
            Assert.NotNull(dn);
            Assert.Equal("$B$2", dn.TextValue);
            Assert.Same(sheet2, dn.LocalSheet);
        }

        [Fact(DisplayName = "Test of multiple defined names preserving insertion order on round-trip")]
        public void DefinedNames_OrderPreserved()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedNameCell("Beta", workbook.CurrentWorksheet, "A1");
            workbook.AddDefinedNameCell("Alpha", workbook.CurrentWorksheet, "A2");
            workbook.AddDefinedNameCell("Gamma", workbook.CurrentWorksheet, "A3");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            System.Collections.Generic.IReadOnlyList<DefinedName> names = given.GetDefinedNames();
            Assert.Equal(3, names.Count);
            Assert.Equal("Beta", names[0].Name);
            Assert.Equal("Alpha", names[1].Name);
            Assert.Equal("Gamma", names[2].Name);
        }

        [Fact(DisplayName = "Test of the comment attribute round-trip on a defined name")]
        public void DefinedNames_CommentRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedNameCell("MyName", workbook.CurrentWorksheet, "A1", null, "this is a comment");
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            DefinedName dn = given.GetDefinedName("MyName");
            Assert.NotNull(dn);
            Assert.Equal("this is a comment", dn.Comment);
        }

        [Theory(DisplayName = "Test of defined name constant round trips")]
        [InlineData("A \"quoted\" value")]
        [InlineData("   ")]
        [InlineData(true)]
        [InlineData(-42)]
        [InlineData(2.5d)]
        public void DefinedNames_ConstantRoundTrip(object value)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedNameConstant("ConstantName", value);
            DefinedName given = TestUtils.WriteAndReadWorkbook(workbook).GetDefinedName("constantname");
            Assert.Equal(DefinedName.NameType.Constant, given.Type);
            Assert.Equal(value, given.Value);
        }

        [Fact(DisplayName = "Test of date and time defined name constant round trips")]
        public void DefinedNames_DateTimeConstantRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AddDefinedNameConstant("DateName", new System.DateTime(2026, 7, 31, 12, 0, 0));
            workbook.AddDefinedNameConstant("TimeName", System.TimeSpan.FromHours(6));
            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.IsType<double>(given.GetDefinedName("DateName").Value);
            Assert.IsType<double>(given.GetDefinedName("TimeName").Value);
        }

        [Fact(DisplayName = "Test of scoped formula reference resolution after round-trip")]
        public void DefinedNames_ScopedFormulaReferencesRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            Worksheet sheet1 = workbook.CurrentWorksheet;
            workbook.AddWorksheet("sheet2");
            Worksheet sheet2 = workbook.CurrentWorksheet;
            DefinedName global = workbook.AddDefinedNameConstant("Rate", 1);
            DefinedName local = workbook.AddDefinedNameConstant("RATE", 2, sheet1);
            sheet1.AddCellReference(local, "A1", (Style)BasicStyles.Bold.Copy(), 2);
            sheet2.AddCellReference(global, "A1", 1);

            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Cell first = given.Worksheets[0].Cells["A1"];
            Cell second = given.Worksheets[1].Cells["A1"];
            Assert.Same(given.GetDefinedName("rate", given.Worksheets[0]), first.Formula.DefinedNameReference);
            Assert.Same(given.GetDefinedName("RATE"), second.Formula.DefinedNameReference);
            Assert.NotNull(first.CellStyle);
            Assert.Equal(Cell.CellType.Number, first.Formula.CachedValueType);
            Assert.Equal("2", first.Formula.CachedValue);
        }

        [Fact(DisplayName = "Test of range formula reference reconstruction after round-trip")]
        public void DefinedNames_RangeFormulaRoundTrip()
        {
            Workbook workbook = new Workbook("sheet1");
            DefinedName range = workbook.AddDefinedNameRange("RangeName", workbook.CurrentWorksheet, "A1:B2");
            workbook.CurrentWorksheet.AddCellReference(range, "D4", 12);

            Workbook given = TestUtils.WriteAndReadWorkbook(workbook);
            Assert.Equal(4, given.CurrentWorksheet.Cells.Count);
            Cell master = given.CurrentWorksheet.Cells["D4"];
            Assert.Same(given.GetDefinedName("rangename"), master.Formula.DefinedNameReference);
            Assert.Equal("D4:E5", master.Formula.FormulaRange);
            Assert.Equal("12", master.Formula.CachedValue);
            Assert.Equal("D4", given.CurrentWorksheet.Cells["E5"].Formula.MasterCellAddress);
        }

        [Fact(DisplayName = "Test of generated defined name workbook XML")]
        public void DefinedNames_WorkbookXmlTest()
        {
            Workbook workbook = new Workbook("Owner's Sheet");
            Worksheet worksheet = workbook.CurrentWorksheet;
            workbook.AddDefinedNameCell("CellName", worksheet, "A1", worksheet, "a & b");
            workbook.AddDefinedNameConstant("StringName", "A \"quote\"");
            string xml = GetZipEntry(workbook, "xl/workbook.xml");
            Assert.Contains("'Owner''s Sheet'!$A$1", xml);
            Assert.Contains("localSheetId=\"0\"", xml);
            Assert.Contains("comment=\"a &amp; b\"", xml);
            Assert.Contains("\"A \"\"quote\"\"\"", xml);
        }

        [Fact(DisplayName = "Test that the compatibility check rejects a read workbook with an external defined-name link")]
        public void DefinedNames_ExternalLinkCompatibilityCheckFailTest()
        {
            // Hypothetical scenario (should normally not be performed manually)
            using (Stream input = TestUtils.GetResource("external_link.xlsx"))
            {
                Workbook workbook = Extensions.WorkbookReader.Load(input);
                DefinedName definedName = workbook.GetDefinedName("externalDefinedName");
                Assert.NotNull(definedName);
                Assert.True(definedName.HasExternalReferences, definedName.TextValue);
                workbook.GetWorksheet("DirectLink").AddCell("local value", "A1");

                CompatibilityProcessor processor = new CompatibilityProcessor();
                processor.Init(new XlsxWriter(workbook), null);
                Assert.Throws<NotSupportedContentException>(() => processor.Execute());
            }
        }

        [Fact(DisplayName = "Test that saving a read workbook with an external defined-name link fails")]
        public void DefinedNames_ExternalLinkWriteFailTest()
        {
            using (Stream input = TestUtils.GetResource("external_link.xlsx"))
            using (MemoryStream output = new MemoryStream())
            {
                Workbook workbook = Extensions.WorkbookReader.Load(input);
                workbook.GetWorksheet("DirectLink").AddCell("local value", "A1");
                Exceptions.IOException exception = Assert.Throws<Exceptions.IOException>(() => workbook.SaveAsStream(output, true));
                Assert.IsType<NotSupportedContentException>(exception.InnerException);
            }
        }

        [Theory(DisplayName = "Test that saving a newly added external defined-name formula fails")]
        [InlineData("SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)")]
        [InlineData("[1]ExternalSheet!$A$1")]
        public void DefinedNames_AddedExternalLinkWriteFailTest(string expression)
        {
            Workbook workbook = new Workbook("sheet1");
            DefinedName definedName = workbook.AddDefinedNameFormula("ExternalFormula", expression);
            Assert.True(definedName.HasExternalReferences);

            using (MemoryStream output = new MemoryStream())
            {
                Exceptions.IOException exception = Assert.Throws<Exceptions.IOException>(() => workbook.SaveAsStream(output, true));
                Assert.IsType<NotSupportedContentException>(exception.InnerException);
            }
        }

        [Fact(DisplayName = "Test that the compatibility check rejects an imported external cell formula")]
        public void Formulas_ImportedExternalLinkCompatibilityCheckFailTest()
        {
            using (Stream input = TestUtils.GetResource("external_link.xlsx"))
            {
                Workbook workbook = Extensions.WorkbookReader.Load(input);
                workbook.RemoveDefinedName("externalDefinedName");
                FormulaData formula = workbook.GetWorksheet("DirectLink").Cells["A1"].Formula;
                Assert.NotNull(formula);
                Assert.True(formula.HasExternalReferences);

                XlsxWriter writer = CreateWriterWithProcessingData(workbook);
                CompatibilityProcessor processor = new CompatibilityProcessor();
                processor.Init(writer, null);
                Assert.Throws<NotSupportedContentException>(() => processor.Execute());
                Assert.True(writer.WriterProcessingData.HasExternalFormulaReferences);
            }
        }

        [Fact(DisplayName = "Test that saving an imported external cell formula fails")]
        public void Formulas_ImportedExternalLinkWriteFailTest()
        {
            using (Stream input = TestUtils.GetResource("external_link.xlsx"))
            using (MemoryStream output = new MemoryStream())
            {
                Workbook workbook = Extensions.WorkbookReader.Load(input);
                workbook.RemoveDefinedName("externalDefinedName");
                Exceptions.IOException exception = Assert.Throws<Exceptions.IOException>(() => workbook.SaveAsStream(output, true));
                Assert.IsType<NotSupportedContentException>(exception.InnerException);
            }
        }

        [Theory(DisplayName = "Test that clearing an imported external cell formula permits saving")]
        [InlineData(true)]
        [InlineData(false)]
        public void Formulas_ImportedExternalLinkClearedTest(bool removeCell)
        {
            using (Stream input = TestUtils.GetResource("external_link.xlsx"))
            using (MemoryStream output = new MemoryStream())
            {
                Workbook workbook = Extensions.WorkbookReader.Load(input);
                workbook.RemoveDefinedName("externalDefinedName");
                Worksheet worksheet = workbook.GetWorksheet("DirectLink");
                if (removeCell)
                {
                    worksheet.RemoveCell("A1");
                }
                else
                {
                    worksheet.AddCell("local value", "A1");
                }

                workbook.SaveAsStream(output, true);
                Assert.True(output.Length > 0);
            }
        }

        [Fact(DisplayName = "Test that an external formula added through the worksheet API is rejected")]
        public void Formulas_AddedExternalLinkCompatibilityCheckFailTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");

            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(CreateWriterWithProcessingData(workbook), null);

            Assert.Throws<NotSupportedContentException>(() => processor.Execute());
        }

        [Fact(DisplayName = "Test that a plugin-provided false external-formula cache is authoritative")]
        public void Formulas_ExternalLinkFalseCacheTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");
            XlsxWriter writer = CreateWriterWithProcessingData(workbook);
            writer.WriterProcessingData.HasExternalFormulaReferences = false;
            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(writer, null);

            processor.Execute();

            Assert.False(writer.WriterProcessingData.HasExternalFormulaReferences);
        }

        [Fact(DisplayName = "Test that an external formula in a copied worksheet is rejected")]
        public void Formulas_CopiedWorksheetExternalLinkCompatibilityCheckFailTest()
        {
            Workbook source = new Workbook("source");
            source.CurrentWorksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");
            Workbook target = new Workbook("target");
            Workbook.CopyWorksheetTo(source.CurrentWorksheet, "copy", target);

            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(CreateWriterWithProcessingData(target), null);

            Assert.Throws<NotSupportedContentException>(() => processor.Execute());
        }

        [Fact(DisplayName = "Test that an external formula in a copied cell is rejected")]
        public void Formulas_CopiedCellExternalLinkCompatibilityCheckFailTest()
        {
            Workbook source = new Workbook("source");
            source.CurrentWorksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");
            Workbook target = new Workbook("target");
            target.CurrentWorksheet.AddCell(source.CurrentWorksheet.Cells["A1"].Copy(), "B2");

            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(CreateWriterWithProcessingData(target), null);

            Assert.Throws<NotSupportedContentException>(() => processor.Execute());
        }

        [Fact(DisplayName = "Test that an external formula stored as a raw cell value is rejected")]
        public void Formulas_RawCellExternalLinkCompatibilityCheckFailTest()
        {
            Workbook workbook = new Workbook("sheet1");
            Cell cell = new Cell("local value", Cell.CellType.String);
            workbook.CurrentWorksheet.AddCell(cell, "A1");
            cell.DataType = Cell.CellType.Formula;
            cell.Value = "[1]ExternalSheet!A1";
            Assert.Null(workbook.CurrentWorksheet.Cells["A1"].Formula);

            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(CreateWriterWithProcessingData(workbook), null);

            Assert.Throws<NotSupportedContentException>(() => processor.Execute());
        }

        [Fact(DisplayName = "Test that a compatible raw formula cell passes external link validation")]
        public void Formulas_RawCellCompatibleExpressionTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(new Cell("Table1[Column]", Cell.CellType.Formula), "A1");
            XlsxWriter writer = CreateWriterWithProcessingData(workbook);
            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(writer, null);

            processor.Execute();

            Assert.False(writer.WriterProcessingData.HasExternalFormulaReferences);
        }

        [Fact(DisplayName = "Test that an external formula stored as a raw non-string value is rejected")]
        public void Formulas_RawObjectExternalLinkCompatibilityCheckFailTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(new Cell(new System.Text.StringBuilder("[1]ExternalSheet!A1"), Cell.CellType.Formula), "A1");
            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(CreateWriterWithProcessingData(workbook), null);

            Assert.Throws<NotSupportedContentException>(() => processor.Execute());
        }

        [Fact(DisplayName = "Test that a null raw formula value passes validation without processing data")]
        public void Formulas_RawNullCompatibleExpressionTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCell(new Cell(null, Cell.CellType.Formula), "A1");
            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(new XlsxWriter(workbook), null);

            processor.Execute();
        }

        [Theory(DisplayName = "Test that compatible formula expressions pass external link validation")]
        [InlineData("SUM(A1:A2)")]
        [InlineData("Table1[Column]")]
        [InlineData("Table1[1]")]
        [InlineData("R[1]C[1]")]
        [InlineData("INDIRECT(\"[1]Sheet1!A1\")")]
        public void Formulas_CompatibleExpressionTest(string expression)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCellFormula(expression, "A1");
            XlsxWriter writer = CreateWriterWithProcessingData(workbook);
            CompatibilityProcessor processor = new CompatibilityProcessor();
            processor.Init(writer, null);

            processor.Execute();
            Assert.False(writer.WriterProcessingData.HasExternalFormulaReferences);
            processor.Execute();
        }

        [Fact(DisplayName = "Test of DefinedNameDefinition properties")]
        public void DefinedNames_DefinitionPropertiesTest()
        {
            // Hypothetical scenario (Is normally only created by a Reader plug-in)
            DefinedNameDefinition definition = new DefinedNameDefinition
            {
                Name = "Name",
                Reference = "1",
                LocalSheetId = 2,
                Comment = "comment"
            };
            Assert.Equal("Name", definition.Name);
            Assert.Equal("1", definition.Reference);
            Assert.Equal(2, definition.LocalSheetId);
            Assert.Equal("comment", definition.Comment);
        }

        [Fact(DisplayName = "Test that FinalizingProcessor defers until all worksheets are available")]
        public void DefinedNames_FinalizingProcessorDeferredTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity, 0, new WorksheetDefinition(1, "sheet1", "rId1"));
            workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity, 1, new WorksheetDefinition(2, "sheet2", "rId2"));
            bool invoked = false;
            FinalizingProcessor processor = new FinalizingProcessor();
            processor.Init(workbook, null, (w, id, options, index) => invoked = id == PlugInUUID.FinalizingInlineProcessor);
            processor.Execute();
            Assert.Empty(workbook.GetDefinedNames());
            Assert.True(invoked);
            Assert.Same(workbook, processor.Workbook);
            Assert.Null(processor.Options);
            Assert.NotNull(processor.InlinePluginHandler);
        }

        [Fact(DisplayName = "Test that FinalizingProcessor handles invalid local sheet IDs")]
        public void DefinedNames_FinalizingProcessorInvalidScopeTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.DefinedNameEntity, 0,
                new DefinedNameDefinition { Name = "Name", Reference = "1", LocalSheetId = 99 });
            FinalizingProcessor processor = new FinalizingProcessor();
            processor.Init(workbook, null, null);
            processor.Execute();
            Assert.Null(workbook.GetDefinedName("Name").LocalSheet);
        }

        [Fact(DisplayName = "Test that FinalizingProcessor skips formula cells consumed by an array")]
        public void DefinedNames_FinalizingProcessorOverlappingArrayTest()
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.CurrentWorksheet.AddCellFormula("RangeName", "D4");
            workbook.CurrentWorksheet.AddCellFormula("RangeName", "E4");
            workbook.AuxiliaryData.SetData(PlugInUUID.WorkbookReader, PlugInUUID.DefinedNameEntity, 0,
                new DefinedNameDefinition { Name = "RangeName", Reference = "'sheet1'!$A$1:$B$1" });
            FinalizingProcessor processor = new FinalizingProcessor();
            processor.Init(workbook, null, null);
            processor.Execute();
            Assert.Equal("D4", workbook.CurrentWorksheet.Cells["E4"].Formula.MasterCellAddress);
        }

        [Fact(DisplayName = "Test that ReadDefinedNameReference returns empty string for a self-closing definedName element")]
        public void DefinedNames_ReadReference_SelfClosing()
        {
            const string xml = "<definedName name=\"X\"/>";
            using (System.IO.StringReader sr = new System.IO.StringReader(xml))
            using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(sr))
            {
                reader.MoveToContent();
                string result = NanoXLSX.Internal.Readers.WorkbookReader.ReadDefinedNameReference(reader);
                Assert.Equal(string.Empty, result);
            }
        }

        [Theory(DisplayName = "Test that ReadDefinedNameReference returns supported text content")]
        [InlineData("<definedName name=\"X\">sheet1!$A$1</definedName>", "sheet1!$A$1")]
        [InlineData("<definedName name=\"X\"><![CDATA[A&B]]></definedName>", "A&B")]
        [InlineData("<definedName name=\"X\" xml:space=\"preserve\">  </definedName>", "  ")]
        public void DefinedNames_ReadReference_TextContent(string xml, string expected)
        {
            using (System.IO.StringReader sr = new System.IO.StringReader(xml))
            using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(sr))
            {
                reader.MoveToContent();
                string result = NanoXLSX.Internal.Readers.WorkbookReader.ReadDefinedNameReference(reader);
                Assert.Equal(expected, result);
            }
        }

        [Theory(DisplayName = "Test of the 'WorkbookProtectionPasswordHash' property when writing and reading a workbook, using legacy password")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("A")]
        [InlineData("123")]
        [InlineData("test")]
        public void ReadWorkbookPasswordHashTest(string plainText)
        {
            Workbook workbook = new Workbook("sheet1");
            workbook.SetWorkbookProtection(true, false, true, plainText);
            Workbook givenWorkbook = TestUtils.WriteAndReadWorkbook(workbook);
            string hash = LegacyPassword.GenerateLegacyPasswordHash(plainText);
            if (hash == "")
            {
                hash = null;
            }
            Assert.Equal(hash, givenWorkbook.WorkbookProtectionPassword.PasswordHash);
        }

        private static string GetZipEntry(Workbook workbook, string entryName)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                workbook.SaveAsStream(stream, true);
                stream.Position = 0;
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
                using (StreamReader reader = new StreamReader(archive.GetEntry(entryName).Open()))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static XlsxWriter CreateWriterWithProcessingData(Workbook workbook)
        {
            return new XlsxWriter(workbook)
            {
                WriterProcessingData = new WriterProcessingData(workbook, StyleRepository.Instance)
            };
        }
    }
}
