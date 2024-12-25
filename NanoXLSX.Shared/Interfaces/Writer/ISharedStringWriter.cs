using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Shared.Interfaces;

namespace NanoXLSX.Interfaces.Writer
{
    public interface ISharedStringWriter : IPluginWriter
    {
        ISortedMap SharedStrings { get; }

        int SharedStringsTotalCount { get; set; }
    }
}
