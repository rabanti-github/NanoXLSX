using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace NanoXLSX.Interfaces.Reader
{
    internal interface IInlinePlugInReader : IPlugIn
    {
        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        Workbook Workbook { get; set; }

        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="stream">Stream, containing the XML file to red</param>
        /// <param name="workbook">Workbook instance where read data is placed</param>
        void Init(ref MemoryStream stream, Workbook workbook);

    }
}
