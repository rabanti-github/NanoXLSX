/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace NanoXLSX.Test.Worksheets
{
    public class CellValuesTest
    {
        [Fact(DisplayName = "CellValues contains the cell added via AddCell")]
        public void CellValues_ContainsAddedCell()
        {
            Worksheet ws = new Worksheet();
            ws.AddCell("hello", 2, 3);
            List<Cell> values = ws.CellValues.ToList();
            Assert.Single(values);
            Assert.Equal("hello", values[0].Value);
        }

        [Fact(DisplayName = "CellValues.Count matches Cells.Count")]
        public void CellValues_CountMatchesCells()
        {
            Worksheet ws = new Worksheet();
            ws.AddCell(1, 0, 0);
            ws.AddCell(2, 1, 0);
            ws.AddCell(3, 2, 0);
            Assert.Equal(ws.Cells.Count, ws.CellValues.Count());
        }

        [Fact(DisplayName = "CellValues and Cells.Values yield the same Cell instances")]
        public void CellValues_SameInstancesAsCells()
        {
            Worksheet ws = new Worksheet();
            ws.AddCell("a", 0, 0);
            ws.AddCell("b", 1, 0);
            ws.AddCell("c", 0, 1);
            HashSet<Cell> fromCellValues = new HashSet<Cell>(ws.CellValues);
            HashSet<Cell> fromCells = new HashSet<Cell>(ws.Cells.Values);
            Assert.Equal(fromCells, fromCellValues);
        }

        [Fact(DisplayName = "Cells[string] resolves the correct cell after AddCell")]
        public void Cells_StringIndexer_ResolvesCorrectCell()
        {
            Worksheet ws = new Worksheet();
            ws.AddCell(42, 0, 0);
            ws.AddCell("world", 3, 5);
            Assert.Equal(42, ws.Cells["A1"].Value);
            Assert.Equal("world", ws.Cells["D6"].Value);
        }

        [Fact(DisplayName = "CellValues is empty after RemoveCell")]
        public void CellValues_EmptyAfterRemove()
        {
            Worksheet ws = new Worksheet();
            ws.AddCell(99, 0, 0);
            ws.RemoveCell(0, 0);
            Assert.Empty(ws.CellValues);
        }

        [Fact(DisplayName = "Cells.ContainsKey returns true for added cell, false after removal")]
        public void Cells_ContainsKey_TracksMutations()
        {
            Worksheet ws = new Worksheet();
            ws.AddCell(7, 1, 2);
            Assert.True(ws.Cells.ContainsKey("B3"));
            ws.RemoveCell(1, 2);
            Assert.False(ws.Cells.ContainsKey("B3"));
        }

        [Fact(DisplayName = "Cell formula type and value changes update worksheet and workbook features")]
        public void CellFormulaChanges_UpdateFeatures()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCell("Initial Value", "A1");
            Cell cell = worksheet.Cells["A1"];

            Assert.Equal(0, worksheet.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);

            cell.DataType = Cell.CellType.Formula;

            Assert.Equal(1, worksheet.Features.FormulaCount);
            Assert.Equal(1, workbook.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);

            cell.Value = "[1]ExternalSheet!A1";

            Assert.Equal(1, worksheet.Features.ExternalLinkCount);
            Assert.Equal(1, workbook.Features.ExternalLinkCount);

            cell.Value = "A1";

            Assert.Equal(0, worksheet.Features.ExternalLinkCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);

            cell.DataType = Cell.CellType.String;

            Assert.Null(cell.Formula);
            Assert.Equal(0, worksheet.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.FormulaCount);
        }

        [Theory(DisplayName = "Changing a formula cell to a non-formula type removes its feature contribution")]
        [InlineData(Cell.CellType.String)]
        [InlineData(Cell.CellType.Number)]
        [InlineData(Cell.CellType.Bool)]
        public void CellFormulaTypeChange_RemovesFeatures(Cell.CellType targetType)
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");
            Cell cell = worksheet.Cells["A1"];

            Assert.Equal(1, workbook.Features.FormulaCount);
            Assert.Equal(1, workbook.Features.ExternalLinkCount);

            cell.DataType = targetType;

            Assert.Equal(targetType, cell.DataType);
            Assert.Null(cell.Formula);
            Assert.Equal(0, worksheet.Features.FormulaCount);
            Assert.Equal(0, worksheet.Features.ExternalLinkCount);
            Assert.Equal(0, workbook.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);
        }

        [Fact(DisplayName = "Replacing formula metadata keeps feature counters balanced")]
        public void CellFormulaReplacement_UpdatesFeatures()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCellFormula("A1", "A1");
            Cell cell = worksheet.Cells["A1"];

            cell.Formula = new FormulaData("[1]ExternalSheet!A1");

            Assert.Equal(1, worksheet.Features.FormulaCount);
            Assert.Equal(1, workbook.Features.FormulaCount);
            Assert.Equal(1, worksheet.Features.ExternalLinkCount);
            Assert.Equal(1, workbook.Features.ExternalLinkCount);

            cell.Value = null;

            Assert.Equal(Cell.CellType.Empty, cell.DataType);
            Assert.Null(cell.Formula);
            Assert.Equal(0, worksheet.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);
        }

        [Fact(DisplayName = "Changing a defined-name formula value removes its resolved reference feature")]
        public void CellDefinedNameFormulaValueChange_UpdatesFeatures()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            DefinedName definedName = workbook.AddDefinedNameConstant("NamedValue", 1);
            worksheet.AddCellReference(definedName, "A1");
            Cell cell = worksheet.Cells["A1"];

            Assert.Equal(1, worksheet.Features.DefinedNameFormulaCount);
            Assert.Equal(1, workbook.Features.DefinedNameFormulaCount);

            cell.Value = "A1";

            Assert.Null(cell.Formula.DefinedNameReference);
            Assert.Equal("A1", cell.Formula.Expression);
            Assert.Equal(0, worksheet.Features.DefinedNameFormulaCount);
            Assert.Equal(0, workbook.Features.DefinedNameFormulaCount);
            Assert.Equal(1, workbook.Features.DefinedNameCount);
            Assert.Equal(1, workbook.Features.FormulaCount);
        }
    }
}
