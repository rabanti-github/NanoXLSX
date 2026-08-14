using System;
using NanoXLSX.Enums;
using NanoXLSX.Exceptions;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Test.Core.WorkbookTest
{
    public class DefinedNameTest
    {
        [Theory(DisplayName = "Test of valid defined name identifiers")]
        [InlineData("Revenue_2026")]
        [InlineData("_private.name")]
        [InlineData("\\legacy")]
        [InlineData("Übersicht")]
        [InlineData("工作表Name")]
        public void ValidNameTest(string name)
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Equal(name, workbook.AddDefinedNameConstant(name, 1).Name);
        }

        [Theory(DisplayName = "Test of invalid defined name identifiers")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("1Name")]
        [InlineData(".Name")]
        [InlineData("Bad Name")]
        [InlineData("Bad-Name")]
        [InlineData("C")]
        [InlineData("r")]
        [InlineData("A1")]
        [InlineData("XFD1048576")]
        public void InvalidNameTest(string name)
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Throws<FormatException>(() => workbook.AddDefinedNameConstant(name, 1));
        }

        [Fact(DisplayName = "Test of the maximum defined name length")]
        public void NameLengthTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.NotNull(workbook.AddDefinedNameConstant(new string('N', 255), 1));
            Assert.Throws<FormatException>(() => workbook.AddDefinedNameConstant(new string('N', 256), 1));
        }

        [Theory(DisplayName = "Test of AddDefinedNameCell overloads")]
        [InlineData(0)] // address
        [InlineData(1)] // row, column
        [InlineData(2)] // address object
        public void AddDefinedNameCellTest(int overload)
        {
            Workbook workbook = new Workbook("Target");
            Worksheet target = workbook.CurrentWorksheet;
            DefinedName name;
            switch (overload)
            {
                case 0:
                    name = workbook.AddDefinedNameCell("CellName", target, "b2", null, "comment");
                    break;
                case 1:
                    name = workbook.AddDefinedNameCell("CellName", target, 1, 1, null, "comment");
                    break;
                default:
                    name = workbook.AddDefinedNameCell("CellName", target, new Address("B2"), null, "comment"); ;
                    break;
            }

            Assert.Equal(DefinedName.NameType.Cell, name.Type);
            Assert.Equal("$B$2", name.TextValue);
            Assert.Equal(new Address("$B$2"), name.Value);
            Assert.Same(target, name.TargetWorksheet);
            Assert.Null(name.LocalSheet);
            Assert.Equal("comment", name.Comment);
            Assert.Equal(Errors.FormulaError.NoError, name.Error);
            Assert.False(name.HasExternalReferences);
        }

        [Theory(DisplayName = "Test of AddDefinedNameRange overloads")]
        [InlineData(0)] // address string
        [InlineData(1)] // start and end address
        [InlineData(2)] // start/end row and start/end column
        [InlineData(3)] // range object
        public void AddDefinedNameRangeTest(int overload)
        {
            Workbook workbook = new Workbook("Target");
            Worksheet target = workbook.CurrentWorksheet;
            DefinedName name;
            switch (overload)
            {
                case 0:
                    name = workbook.AddDefinedNameRange("RangeName", target, "a1:b3");
                    break;
                case 1:
                    name = workbook.AddDefinedNameRange("RangeName", target, new Address("A1"), new Address("B3"));
                    break;
                case 2:
                    name = workbook.AddDefinedNameRange("RangeName", target, 0, 0, 1, 2);
                    break;
                default:
                    name = workbook.AddDefinedNameRange("RangeName", target, new Range("A1:B3"));
                    break;
            }
            Assert.Equal(DefinedName.NameType.Range, name.Type);
            Assert.Equal("$A$1:$B$3", name.TextValue);
            Assert.Equal(new Range("$A$1:$B$3"), name.Value);
            Assert.Same(target, name.TargetWorksheet);
        }

        [Theory(DisplayName = "Test of all supported defined name constant types")]
        [InlineData("String")]
        [InlineData("Whitespace")]
        [InlineData("Bool")]
        [InlineData("Byte")]
        [InlineData("SByte")]
        [InlineData("Decimal")]
        [InlineData("Double")]
        [InlineData("Float")]
        [InlineData("Int")]
        [InlineData("UInt")]
        [InlineData("Long")]
        [InlineData("ULong")]
        [InlineData("Short")]
        [InlineData("UShort")]
        [InlineData("DateTime")]
        [InlineData("TimeSpan")]
        [InlineData("Object")]
        public void AddDefinedNameConstantTest(string kind)
        {
            object value = CreateConstant(kind);
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name = workbook.AddDefinedNameConstant("ConstantName", value);

            Assert.Equal(DefinedName.NameType.Constant, name.Type);
            Assert.Same(value, name.Value);
            Assert.NotNull(name.TextValue);
            Assert.Null(name.TargetWorksheet);
        }

        [Theory(DisplayName = "Test of invalid defined name values (formula)")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData(" ")]
        public void InvalidFormulaTest(string formula)
        {
            Workbook workbook = new Workbook("Sheet1");
            Assert.Throws<WorksheetException>(() => workbook.AddDefinedNameFormula("FormulaName", formula));
        }

        [Fact(DisplayName = "Test of formula and null constant defined names")]
        public void FormulaAndNullConstantTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName formula = workbook.AddDefinedNameFormula("FormulaName", "SUM(1,2)", null, "note");
            Assert.Equal(DefinedName.NameType.Formula, formula.Type);
            Assert.Equal("SUM(1,2)", formula.TextValue);
            Assert.Equal("note", formula.Comment);
            Assert.False(formula.HasExternalReferences);
            Assert.Throws<WorksheetException>(() => workbook.AddDefinedNameConstant("NullName", null));
            Assert.Throws<FormatException>(() => workbook.AddDefinedNameConstant("EmptyName", string.Empty));
        }

        [Theory(DisplayName = "Test of external workbook reference detection in added defined-name formulas")]
        [InlineData("[book.xlsx]Sheet1!A1")]
        [InlineData("'..\\[book.xlsx]Data'!$B$2")]
        [InlineData("'../[book.xlsx]Data'!$B$2")]
        [InlineData("'C:\\temp\\[book one.xlsx]Sheet 1'!$A$1")]
        [InlineData("SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1,'..\\[other.xlsx]Data'!$B$2)")]
        [InlineData("[1]Sheet1!$A$1")]
        public void AddDefinedNameFormulaExternalReferenceTest(string expression)
        {
            Workbook workbook = new Workbook("Sheet1");

            DefinedName name = workbook.AddDefinedNameFormula("ExternalFormula", expression);

            Assert.True(name.HasExternalReferences);
            Assert.Equal(expression, name.TextValue);
            Assert.Equal(expression, name.Value);
        }

        [Theory(DisplayName = "Test of non-external defined-name formulas")]
        [InlineData("SUM(A1:A2)")]
        [InlineData("Table1[Column]")]
        [InlineData("R[1]C[1]")]
        [InlineData("INDIRECT(\"[1]Sheet1!A1\")")]
        public void AddDefinedNameFormulaWithoutExternalReferenceTest(string expression)
        {
            Workbook workbook = new Workbook("Sheet1");

            DefinedName name = workbook.AddDefinedNameFormula("LocalFormula", expression);

            Assert.False(name.HasExternalReferences);
        }

        [Fact(DisplayName = "Test of invalid cell and range defined names")]
        public void InvalidCellAndRangeTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            Assert.Throws<WorksheetException>(() => workbook.AddDefinedNameCell("Name1", null, new Address("A1")));
            Assert.Throws<FormatException>(() => workbook.AddDefinedNameCell("Name2", worksheet, (Address)null));
            Assert.Throws<WorksheetException>(() => workbook.AddDefinedNameRange("Name3", null, new Range("A1:B2")));
            Assert.Throws<FormatException>(() => workbook.AddDefinedNameRange("Name4", worksheet, (Range)null));
            Assert.Throws<FormatException>(() => workbook.AddDefinedNameCell("Name5", worksheet, "A1:B2"));
            Assert.Equal("$A$1:$A$1", workbook.AddDefinedNameRange("Name6", worksheet, "A1").TextValue);
        }

        [Fact(DisplayName = "Test of case-insensitive identity and defined name scopes")]
        public void ScopeAndIdentityTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet sheet1 = workbook.CurrentWorksheet;
            workbook.AddWorksheet("Sheet2");
            Worksheet sheet2 = workbook.CurrentWorksheet;
            DefinedName global = workbook.AddDefinedNameConstant("Rate", 1);
            DefinedName local1 = workbook.AddDefinedNameConstant("Rate", 2, sheet1);
            DefinedName local2 = workbook.AddDefinedNameConstant("RATE", 3, sheet2);

            Assert.Same(global, workbook.GetDefinedName("rate"));
            Assert.Same(local1, workbook.GetDefinedName("RATE", sheet1));
            Assert.Same(local2, workbook.GetDefinedName("rate", sheet2));
            Assert.Null(workbook.GetDefinedName("missing"));
            Assert.Throws<WorksheetException>(() => workbook.AddDefinedNameConstant("rAtE", 4));
            Assert.Throws<WorksheetException>(() => workbook.AddDefinedNameConstant("rAtE", 4, sheet1));
            Assert.Equal(3, workbook.GetDefinedNames().Count);
        }

        [Fact(DisplayName = "Test of removing a defined name and invalidating formula references")]
        public void RemoveDefinedNameTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName removed = workbook.AddDefinedNameConstant("RemovedName", 5);
            DefinedName retained = workbook.AddDefinedNameConstant("RetainedName", 6);
            workbook.CurrentWorksheet.AddCellReference(removed, "A1");
            workbook.CurrentWorksheet.AddCellReference(retained, "A2");
            workbook.AddWorksheet("Sheet2");
            workbook.CurrentWorksheet.AddCellReference(removed, "B1");

            Assert.True(workbook.RemoveDefinedName("removedname"));
            Assert.Null(workbook.Worksheets[0].Cells["A1"].Formula.DefinedNameReference);
            Assert.Null(workbook.Worksheets[1].Cells["B1"].Formula.DefinedNameReference);
            Assert.Same(retained, workbook.Worksheets[0].Cells["A2"].Formula.DefinedNameReference);
            Assert.False(workbook.RemoveDefinedName("missing"));
        }

        [Fact(DisplayName = "Test of removing all defined names and invalidating formula references")]
        public void RemoveAllDefinedNameTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name1 = workbook.AddDefinedNameConstant("name1", 5);
            DefinedName name2 = workbook.AddDefinedNameConstant("name2", 6);
            workbook.CurrentWorksheet.AddCellReference(name1, "A1");
            workbook.CurrentWorksheet.AddCellReference(name2, "A2");
            workbook.AddWorksheet("Sheet2");
            workbook.CurrentWorksheet.AddCellReference(name1, "B1");

            workbook.ClearDefinedNames();

            Assert.Null(workbook.Worksheets[0].Cells["A1"].Formula.DefinedNameReference);
            Assert.Null(workbook.Worksheets[0].Cells["A2"].Formula.DefinedNameReference);
            Assert.Null(workbook.Worksheets[1].Cells["B1"].Formula.DefinedNameReference);
            Assert.Empty(workbook.GetDefinedNames());
            Assert.NotNull(workbook.GetDefinedNames());
        }

        [Fact(DisplayName = "Test of DefinedName equality, hashing and object equality")]
        public void EqualityTest()
        {
            Workbook workbook1 = new Workbook("Sheet1");
            Workbook workbook2 = new Workbook("Sheet1");
            DefinedName a = workbook1.AddDefinedNameConstant("CaseName", 1, null, "comment");
            DefinedName b = workbook2.AddDefinedNameConstant("casename", 1, null, "comment");

            Assert.True(a.Equals(a));
            Assert.True(a.Equals(b));
            Assert.True(a.Equals((object)b));
            Assert.Equal(a.GetHashCode(), b.GetHashCode());
            Assert.False(a.Equals((DefinedName)null));
            Assert.False(a.Equals((object)null));
            Assert.False(a.Equals("wrong"));
            Assert.False(a.Equals(workbook2.AddDefinedNameConstant("OtherName", 1)));
        }

        [Fact(DisplayName = "Test of DefinedName comparison and string representation")]
        public void CompareToAndToStringTest()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet sheet = workbook.CurrentWorksheet;
            DefinedName a = workbook.AddDefinedNameConstant("Alpha", 1);
            DefinedName sameNameLocal = workbook.AddDefinedNameConstant("alpha", 1, sheet);
            DefinedName z = workbook.AddDefinedNameFormula("Zulu", "1+1");

            Assert.Equal(1, a.CompareTo(null));
            Assert.True(a.CompareTo(z) < 0);
            Assert.True(a.CompareTo(sameNameLocal) < 0);
            Assert.True(sameNameLocal.CompareTo(a) > 0);
            Assert.Contains("name=Alpha", a.ToString());
            Assert.Contains("scope=workbook", a.ToString());
            Assert.Contains("sheet:Sheet1", sameNameLocal.ToString());
        }

        [Fact(DisplayName = "Test of every DefinedName comparison component")]
        public void CompareToComponentsTest()
        {
            Workbook workbook1 = new Workbook("Sheet1");
            workbook1.AddWorksheet("Sheet2");
            Workbook workbook2 = new Workbook("Other");
            Worksheet first = workbook1.Worksheets[0];
            Worksheet second = workbook1.Worksheets[1];
            DefinedName baseline = new DefinedName(workbook1, DefinedName.NameType.Constant, "Name", 1, null, first, "a");

            Assert.NotEqual(0, baseline.CompareTo(new DefinedName(workbook2, DefinedName.NameType.Formula, "Name", "1", null, first, "a")));
            Assert.True(baseline.CompareTo(new DefinedName(workbook2, DefinedName.NameType.Constant, "Name", 1, null, second, "a")) < 0);
            Assert.Equal(0, baseline.CompareTo(new DefinedName(workbook2, DefinedName.NameType.Constant, "name", 1, null, first, "a")));
            Assert.True(baseline.CompareTo(new DefinedName(workbook2, DefinedName.NameType.Constant, "Name", 2, null, first, "a")) < 0);
            Assert.True(baseline.CompareTo(new DefinedName(workbook2, DefinedName.NameType.Constant, "Name", 1, null, first, "b")) < 0);

            DefinedName target1 = new DefinedName(workbook1, DefinedName.NameType.Cell, "Target", "A1", first, null);
            DefinedName target2 = new DefinedName(workbook2, DefinedName.NameType.Cell, "Target", "A1", second, null);
            Assert.True(target1.CompareTo(target2) < 0);
        }

        [Theory(DisplayName = "Test of resolving defined names read from workbook XML")]
        [InlineData("\"text\"", DefinedName.NameType.Constant, "text")]
        [InlineData("TRUE", DefinedName.NameType.Constant, "TRUE")]
        [InlineData("42", DefinedName.NameType.Constant, "42")]
        [InlineData("2.5", DefinedName.NameType.Constant, "2.5")]
        [InlineData("'Sheet1'!$A$1", DefinedName.NameType.Cell, "$A$1")]
        [InlineData("'Sheet1'!$A$1:$B$2", DefinedName.NameType.Range, "$A$1:$B$2")]
        [InlineData("SUM(1,2)", DefinedName.NameType.Formula, "SUM(1,2)")]
        [InlineData("#REF!", DefinedName.NameType.Formula, "#REF!")]
        [InlineData("'Missing'!$A$1", DefinedName.NameType.Cell, "$A$1")]
        [InlineData("'Sheet1'!invalid", DefinedName.NameType.Formula, "'Sheet1'!invalid")]
        public void ResolveDefinedNameTest(string reference, DefinedName.NameType expectedType, string expectedText)
        {
            Workbook workbook = new Workbook("Sheet1");
            DefinedName name = DefinedName.ResolveDefinedName("ResolvedName", reference, workbook, null, "comment");
            Assert.Equal(expectedType, name.Type);
            Assert.Equal(expectedText, name.TextValue);
            Assert.Equal("comment", name.Comment);
            Assert.Equal(reference == "#REF!" ? Errors.FormulaError.Reference : Errors.FormulaError.NoError, name.Error);
        }

        [Theory(DisplayName = "Test external references in resolved defined names")]
        [InlineData("'[1]Sheet1'!$A$1")]
        [InlineData("SUM([1]Sheet1!$A$1)")]
        [InlineData("SUM('C:\\temp\\[book one.xlsx]Sheet 1'!$A$1)")]
        public void ResolveDefinedNameExternalReferenceTest(string reference)
        {
            Workbook workbook = new Workbook("Sheet1");

            DefinedName name = DefinedName.ResolveDefinedName("ResolvedExternalName", reference, workbook, null, null);

            Assert.True(name.HasExternalReferences);
        }

        [Fact(DisplayName = "Test that DefinedName requires a workbook")]
        public void NullWorkbookTest()
        {
            Assert.Throws<FormatException>(() => new DefinedName(null, DefinedName.NameType.Constant, "Name", 1, null));
            Assert.Throws<FormatException>(() => new DefinedName(new Workbook(), DefinedName.NameType.Formula, "Name", "", null));
        }

        [Theory(DisplayName = "Test of the ReplaceExpression method")]
        [InlineData("$A$1-$A$2", "$B$2-$C$2")]
        [InlineData("A1", "B1")]
        [InlineData("x", "x")] // keep
        [InlineData("A", "a")] // case
        [InlineData("[1]worksheet1!$C$1", "[extWorkbook.xlsx]worksheet1!$C$1")]
        public void ReplaceExpressionTest(string oldExpression, string newExpression)
        {
            Workbook wb = new Workbook("sheet1");
            wb.AddDefinedName("name", DefinedName.NameType.Formula, oldExpression, null, null, "comment");
            Assert.Equal(oldExpression, wb.GetDefinedName("name").TextValue);

            wb.GetDefinedName("name").ReplaceExpression(newExpression);

            DefinedName name = wb.GetDefinedName("name");
            Assert.Equal(newExpression, name.TextValue);
            Assert.Equal(DefinedName.NameType.Formula, name.Type);
            Assert.Equal("comment", name.Comment);
        }

        [Theory(DisplayName = "Test of the ignoring ReplaceExpression method on incompatible types")]
        [InlineData(DefinedName.NameType.Constant, "A")]
        [InlineData(DefinedName.NameType.Cell, "$B$2")]
        [InlineData(DefinedName.NameType.Range, "$A$1:$C$2")]
        public void ReplaceExpressionIgnoreTest(DefinedName.NameType type, string expression)
        {
            Workbook wb = new Workbook("sheet1");
            wb.AddDefinedName("name", type, expression, null, null, "comment");
            Assert.Equal(expression, wb.GetDefinedName("name").TextValue);

            wb.GetDefinedName("name").ReplaceExpression("newExpression");

            DefinedName name = wb.GetDefinedName("name");
            Assert.Equal(expression, name.TextValue);
            Assert.Equal(type, name.Type);
            Assert.Equal("comment", name.Comment);
        }

        private static object CreateConstant(string kind)
        {
            switch (kind)
            {
                case "String": return "A \"quoted\" string";
                case "Whitespace": return "   ";
                case "Bool": return true;
                case "Byte": return (byte)1;
                case "SByte": return (sbyte)-2;
                case "Decimal": return 3.25m;
                case "Double": return 4.5d;
                case "Float": return 5.75f;
                case "Int": return -6;
                case "UInt": return (uint)7;
                case "Long": return (long)-8;
                case "ULong": return (ulong)9;
                case "Short": return (short)-10;
                case "UShort": return (ushort)11;
                case "DateTime": return new DateTime(2026, 7, 31, 12, 30, 0);
                case "TimeSpan": return TimeSpan.FromHours(13.5);
                default: return new ConstantObject();
            }
        }

        private sealed class ConstantObject
        {
            public override string ToString()
            {
                return "custom";
            }
        }
    }
}
