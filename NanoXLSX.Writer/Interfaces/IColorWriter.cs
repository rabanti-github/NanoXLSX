/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Colors;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by specific writers that provides color handling
    /// </summary>
    public interface IColorWriter
    {
        /// <summary>
        /// Gets the attribute name for the given color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>Attribute name</returns>
        string GetAttributeName(Color color);
        /// <summary>
        /// Gets the attribute value for the given color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>Attribute value</returns>
        string GetAttributeValue(Color color);

        /// <summary>
        /// Gets whether a tint value is used for the given color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>True if tint is used</returns>
        bool UseTintAttribute(Color color);

        /// <summary>
        /// Gets the tint value as string of the given color instance, if applicable (see <see cref="UseTintAttribute(Color)"/>)
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>Tint value (-1 to 1) as string</returns>
        string GetTintAttributeValue(Color color);

        /// <summary>
        /// Gets all applicable attributes of the given color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>IEnumerable of the applicable XmlAttribute values of the color</returns>
        IEnumerable<XmlAttribute> GetAttributes(Color color);
    }
}
