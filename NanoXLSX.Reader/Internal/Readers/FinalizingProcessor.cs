/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a processor (no stream handling) for finalizing tasks after all parts are read from a XLSX file
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.FinalizingProcessor)]
    public class FinalizingProcessor : IPluginReadProcessor
    {
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }

        /// <summary>
        /// Reference to the <see cref="ReaderPlugInHandler"/>, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }

        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in processor methods</param>
        public void Init(Workbook workbook, IOptions readerOptions, Action<Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            // stream can be ignore
            Workbook = workbook;
            InlinePluginHandler = inlinePluginHandler;
            Options = readerOptions;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        public void Execute()
        {
            FinalizeDefinedNames();

            // TODO Add further regular finalizing tasks here

            InlinePluginHandler?.Invoke(Workbook, PlugInUUID.FinalizingInlineProcessor, Options, null);
        }

        /// <summary>
        /// Finalizes the stashed defined-name definitions into <see cref="DefinedName"/> instances on the workbook,
        /// once all expected worksheets have been bound. Defined-name resolution requires worksheet references for
        /// worksheets (for cell and range references and localSheetId scoping; therefore, this step is deferred until after the last worksheet is wired up.
        /// </summary>
        private void FinalizeDefinedNames()
        {
            List<WorksheetDefinition> worksheetDefinitions = Workbook.AuxiliaryData.GetDataList<WorksheetDefinition>(PlugInUUID.WorkbookReader, PlugInUUID.WorksheetDefinitionEntity);
            if (Workbook.Worksheets.Count < worksheetDefinitions.Count)
            {
                return;
            }
            List<DefinedNameDefinition> definitions = Workbook.AuxiliaryData.GetDataList<DefinedNameDefinition>(PlugInUUID.WorkbookReader, PlugInUUID.DefinedNameEntity);
            foreach (DefinedNameDefinition definition in definitions)
            {
                Worksheet localSheet = null;
                if (definition.LocalSheetId.HasValue)
                {
                    int i = definition.LocalSheetId.Value;
                    if (i >= 0 && i < Workbook.Worksheets.Count)
                    {
                        localSheet = Workbook.Worksheets[i];
                    }
                }
                DefinedName resolvedDefinedName = DefinedName.ResolveDefinedName(definition.Name, definition.Reference, Workbook, localSheet, definition.Comment);
                Workbook.AddDefinedName(resolvedDefinedName);
            }
            RetagReferenceCells(); // TODO check & handle array cell references
        }

        /// <summary>
        /// Tags formula cells whose value matches a workbook defined name as formula with an defined name.
        /// Runs once, after all worksheets are bound and defined names finalized, since both worksheets and
        /// defined names must be in place to perform the lookup.
        /// </summary>
        private void RetagReferenceCells()
        {
            if (!Workbook.Features.ContainsDefinedNames)
            {
                return;
            }
            foreach (Worksheet ws in Workbook.Worksheets)
            {
                Dictionary<string, Tuple<Cell, DefinedName>> referenceCellCopies = new Dictionary<string, Tuple<Cell, DefinedName>>(StringComparer.Ordinal);
                foreach (Cell cell in ws.Cells.Values)
                {
                    if (cell.DataType == Cell.CellType.Formula && cell.Formula != null)
                    {
                        DefinedName definedName = Workbook.GetDefinedName(cell.Formula.Expression, ws)
                            ?? Workbook.GetDefinedName(cell.Formula.Expression);
                        if (definedName != null)
                        {
                            referenceCellCopies.Add(cell.CellAddress, new Tuple<Cell, DefinedName>(cell.Copy(), definedName));
                        }
                    }
                }
                if (referenceCellCopies.Count == 0)
                {
                    continue;
                }
                HashSet<string> processedAddresses = new HashSet<string>(StringComparer.Ordinal);
                foreach (KeyValuePair<string, Tuple<Cell, DefinedName>> cell in referenceCellCopies)
                {
                    if (processedAddresses.Contains(cell.Key))
                    {
                        continue; // Skip processed cells 
                    }
                    IReadOnlyList<Address> addresses;
                    object cachedValue = cell.Value.Item1.Formula.CachedValue;
                    if (cell.Value.Item1.CellStyle == null) // Re-add formula cells with full resolution
                    {
                        addresses = ws.AddCellReference(cell.Value.Item2, cell.Key, cachedValue);
                    }
                    else
                    {
                        addresses = ws.AddCellReference(cell.Value.Item2, cell.Key, cell.Value.Item1.CellStyle, cachedValue);
                    }
                    FormulaData restoredFormula = ws.Cells[cell.Key].Formula;
                    restoredFormula.CachedValue = cachedValue;
                    restoredFormula.CachedValueType = cell.Value.Item1.Formula.CachedValueType;
                    foreach (Address address in addresses)
                    {
                        processedAddresses.Add(address.ToString());
                    }
                }
            }
        }

    }
}
