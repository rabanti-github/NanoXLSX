using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Interfaces.Writer
{
    public interface IPasswordWriter : IPassword
    {
        string GetXmlAttributes();
    }
}
