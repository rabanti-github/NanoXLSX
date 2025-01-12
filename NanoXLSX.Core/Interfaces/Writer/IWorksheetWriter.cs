using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Interfaces.Writer
{
    public interface IWorksheetWriter : IPluginWriter
    {
        Worksheet CurrentWorksheet { get; set; }
    }
}
