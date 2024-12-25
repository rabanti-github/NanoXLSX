using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Interfaces.Workbook;

namespace NanoXLSX.Interfaces.Writer
{
    public interface IWorksheetWriter : IPluginWriter
    {
        IWorksheet CurrentWorksheet { get; set; }
    }
}
