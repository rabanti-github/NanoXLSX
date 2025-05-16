using System.IO;

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
        /// <param name="index">Optional index, e.g. for worksheet identification</param>
        void Init(ref MemoryStream stream, Workbook workbook, int? index = null);

    }
}
