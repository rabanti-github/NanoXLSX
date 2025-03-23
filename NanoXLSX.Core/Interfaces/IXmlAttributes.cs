/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to represent an enumeration of internally used XML attribute for the writer process
    /// </summary>
    public interface IXmlAttributes
    {
        /// <summary>
        /// Gets an IEnumerable of XML attributes
        /// </summary>
        /// <returns></returns>
        IEnumerable<XmlAttribute> GetAttributes();
    }
}
