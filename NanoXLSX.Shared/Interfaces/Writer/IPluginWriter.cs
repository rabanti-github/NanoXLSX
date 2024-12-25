/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Interfaces.Workbook;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by XML writer classes 
    /// </summary>
    public interface IPluginWriter : IPlugin
    {
        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        IWorkbook Workbook { get; set; }
        
        /// <summary>
        /// Next plug-in writer to be executed, if not null
        /// </summary>
        IPluginWriter Next { get; set; }

        /// <summary>
        /// Interface function to write an XML file, as a part of an XLSX file
        /// </summary>
        /// <returns></returns>
        string CreateDocument();

        /// <summary>
        /// Method that is called before the <see cref="CreateDocument()"/> method is executed
        /// </summary>
        /// <param name="workbook">Workbook instance (data source)</param>
        void PreWrite(IWorkbook workbook);

        /// <summary>
        /// Method that is called after the <see cref="CreateDocument()"/> method is executed
        /// </summary>
        /// <param name="workbook">Workbook instance (data source)</param>
        void PostWrite(IWorkbook workbook);

        /// <summary>
        /// Gets the unique class ID. This ID is used to identify the class when replacing functionality by extension packages
        /// </summary>
        /// <returns>GUID of the class</returns>
        string GetClassID();

    }
}
