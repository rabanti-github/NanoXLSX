using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace NanoXLSX.Interfaces.Reader
{
    public interface IPasswordReader : IPassword
    {
        void ReadXmlAttributes(XmlNode node);
    }
}
