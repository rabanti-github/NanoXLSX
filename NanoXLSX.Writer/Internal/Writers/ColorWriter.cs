/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Colors;
using NanoXLSX.Interfaces;
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
        public IEnumerable<XmlAttribute> GetAttributes(Color color)
        {
            List<XmlAttribute> attributes = new List<XmlAttribute>();
            if (color == null)
            {
                return attributes;
            }
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
        public string GetAttributeName(Color color)
        {
            if (color == null)
            {
                return null;
            }
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
        public string GetAttributeValue(Color color)
        {
            if (color == null)
            {
                return null;
            }
            switch (color.Type)
            {
                case Color.ColorType.Auto:
                    return "1";
                case Color.ColorType.Rgb:
                    return color.RgbColor.StringValue;
                case Color.ColorType.Indexed:
                    return color.IndexedColor.StringValue;
                case Color.ColorType.Theme:
                    return color.ThemeColor.StringValue;
                case Color.ColorType.System:
                    return color.SystemColor.StringValue;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Gets the numeric string value of the tint attribute of the specified color
        /// </summary>
        /// <param name="color">Color instance</param>
        /// <returns>String representation of the tint value, or null if not specified</returns>
        public string GetTintAttributeValue(Color color)
        {
            return color == null ? null : ParserUtils.ToString(color.Tint.Value);
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
