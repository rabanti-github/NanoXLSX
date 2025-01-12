/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Utils;

namespace NanoXLSX.Internal.Writers
{
    internal class SharedStringWriter : ISharedStringWriter
    {
        private ISortedMap sharedStrings;
        private int sharedStringsTotalCount;

        public int SharedStringsTotalCount
        {
            get { return sharedStringsTotalCount; }
            set { sharedStringsTotalCount = value; }
        }

        public ISortedMap SharedStrings
        {
            get { return sharedStrings; }
        }

        public SharedStringWriter(XlsxWriter writer)
        {
            this.Workbook = writer.Workbook;
            sharedStrings = new SortedMap();
            sharedStringsTotalCount = 0;
        }

        /// <summary>
        /// Method to create shared strings as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        public string CreateDocument(string currentDocument = null)
        {
            PreWrite(Workbook);
            StringBuilder sb = new StringBuilder();
            sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"");
            sb.Append(ParserUtils.ToString(sharedStringsTotalCount));
            sb.Append("\" uniqueCount=\"");
            sb.Append(ParserUtils.ToString(sharedStrings.Count));
            sb.Append("\">");
            foreach (IFormattableText text in sharedStrings.Keys)
            {
                sb.Append("<si>");
                text.AddFormattedValue(sb);
                sb.Append("</si>");
            }
            sb.Append("</sst>");
            PostWrite(Workbook);
            if (NextWriter != null)
            {
                NextWriter.Workbook = this.Workbook;
                return NextWriter.CreateDocument(sb.ToString());
            }
            else
            {
                return sb.ToString();
            }
        }

        /// <summary>
        /// Method that is called before the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PreWrite(Workbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Method that is called after the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PostWrite(Workbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Gets the unique class ID. This ID is used to identify the class when replacing functionality by extension packages
        /// </summary>
        /// <returns>GUID of the class</returns>
        public string GetClassID()
        {
            return "B65BDF84-90E8-4952-A2DF-E28C769E6062";
        }

        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Gets or sets the next plug-in writer. If not null, the next writer to be applied on the document can be called by this property
        /// </summary>
        public IPluginWriter NextWriter { get; set; } = null;

    }
}
