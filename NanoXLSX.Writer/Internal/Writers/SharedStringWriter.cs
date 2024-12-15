using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Interfaces.Workbook;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Utils;

namespace NanoXLSX.Internal.Writers
{
    internal class SharedStringWriter : IPluginWriter
    {
        private readonly Workbook workbook;

        private SortedMap sharedStrings;
        private int sharedStringsTotalCount;

        public int SharedStringsTotalCount
        {
            get { return sharedStringsTotalCount; }
            set { sharedStringsTotalCount = value; }
        }

        public SortedMap SharedStrings
        {
            get { return sharedStrings; }
        }

        public SharedStringWriter (XlsxWriter writer)
        {
            this.workbook = writer.Workbook;
            sharedStrings = new SortedMap();
            sharedStringsTotalCount = 0;
        }

        /// <summary>
        /// Method to create shared strings as raw XML string
        /// </summary>
        /// <returns>Raw XML string</returns>
        public string CreateDocument()
        {
            PreWrite(workbook);
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
            PostWrite(workbook);
            return sb.ToString();
        }

        /// <summary>
        /// Method that is called before the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PreWrite(IWorkbook workbook)
        {
            // NoOp - replaced by plugin
        }

        /// <summary>
        /// Method that is called after the <see cref="CreateDocument()"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="workbook">Workbook instance that is used in this writer</param>
        public virtual void PostWrite(IWorkbook workbook)
        {
            // NoOp - replaced by plugin
        }
    }
}
