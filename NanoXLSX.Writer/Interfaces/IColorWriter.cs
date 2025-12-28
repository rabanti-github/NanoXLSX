using NanoXLSX.Colors;
using NanoXLSX.Utils.Xml;
using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Interfaces
{
    public interface IColorWriter
    {
        string GetAttributeName(Color color);
        string GetAttributeValue(Color color);

        bool UseTintAttribute(Color color);
        string GetTintAttributeValue(Color color);

        IEnumerable<XmlAttribute> GetAttributes(Color color);


    }
}
