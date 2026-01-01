/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Colors;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using System.Collections.Generic;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class to write color values
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.ColorWriter)]
    internal class ColorWriter : IColorWriter
    {
        /// <summary>
        /// Gets the XML attributes for the specified color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>Enumeration of the attributes, describing the color instance</returns>
        /// \remark <remarks>The passed color instance may never be null. Such cases have to be handled earlier</remarks>
        public IEnumerable<XmlAttribute> GetAttributes(Color color)
        {
            List<XmlAttribute> attributes = new List<XmlAttribute>();
            string name = GetAttributeName(color);
            if (name != null)
            {
                string value = GetAttributeValue(color);
                attributes.Add(new XmlAttribute(name, value));
            }
            if (UseTintAttribute(color))
            {
                attributes.Add(new XmlAttribute("tint", GetTintAttributeValue(color)));
            }
            return attributes;
        }

        /// <summary>
        /// Gets the attribute name for the specified color instance
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>Attribute name</returns>
        /// \remark <remarks>The passed color instance may never be null. Such cases have to be handled earlier</remarks>
        public string GetAttributeName(Color color)
        {
            switch (color.Type)
            {
                case Color.ColorType.Auto:
                    return "auto";
                case Color.ColorType.Rgb:
                    return "rgb";
                case Color.ColorType.Indexed:
                    return "indexed";
                case Color.ColorType.Theme:
                    return "theme";
                case Color.ColorType.System:
                    return "system";
                default:
                    return null;
            }
        }

        /// <summary>
        /// Gets the string value of the specified color for the corresponding attribute
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>Attribute value</returns>
        /// \remark <remarks>The passed color instance may never be null. Such cases have to be handled earlier</remarks>
        public string GetAttributeValue(Color color)
        {
            string value = null;
            switch (color.Type)
            {
                case Color.ColorType.Auto:
                    value = "1";
                    break;
                case Color.ColorType.Rgb:
                    value = color.RgbColor.StringValue;
                    break;
                case Color.ColorType.Indexed:
                    value = color.IndexedColor.StringValue;
                    break;
                case Color.ColorType.Theme:
                    value = color.ThemeColor.StringValue;
                    break;
                case Color.ColorType.System:
                    value = color.SystemColor.StringValue;
                    break;
            }
            return value;
        }

        /// <summary>
        /// Gets the numeric string value of the tint attribute of the specified color
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>String representation of the tint value, or null if not specified</returns>
        public string GetTintAttributeValue(Color color)
        {
            if (!color.Tint.HasValue)
            {
                return null;
            }
            double tint = color.Tint.Value;
            return tint == 0.0 ? null : ParserUtils.ToString(tint);
        }

        /// <summary>
        /// Gets whether the tint attribute is used in the specified color
        /// </summary>
        /// <param name="color">Color to check</param>
        /// <returns>True if tint is used</returns>
        public bool UseTintAttribute(Color color)
        {
            return color != null && color.Type == Color.ColorType.Theme && (!color.Tint.HasValue || !Comparators.IsZero(color.Tint.Value));
        }
    }
}
