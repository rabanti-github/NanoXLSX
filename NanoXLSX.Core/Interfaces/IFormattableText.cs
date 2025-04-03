/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to represent complex text data that can be formatted somehow
    /// </summary>
    public interface IFormattableText
    {
        /// <summary>
        /// Gets the main XML element of the formattable text
        /// </summary>
        XmlElement GetXmlElement();
    }
}
