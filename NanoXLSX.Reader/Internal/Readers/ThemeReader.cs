/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Themes;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for theme definitions of XLSX files.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.ThemeReader)]
    public class ThemeReader : IPlugInReader
    {

        private MemoryStream stream;

        #region properties
        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        internal ThemeReader()
        {
        }
        #endregion

        #region methods
        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="stream">MemoryStream to be read</param>
        /// <param name="workbook">Workbook reference</param>
        /// <param name="readerOptions">Reader options</param>
        public void Init(MemoryStream stream, Workbook workbook, IOptions readerOptions)
        {
            this.stream = stream;
            this.Workbook = workbook;
        }

        /// <summary>
        /// Method to execute the main logic of the plug-in (interface implementation)
        /// </summary>
        /// <exception cref="Exceptions.IOException">Throws an IOException in case of a error during reading</exception>
        public void Execute()
        {
            try
            {
                using (stream) // Close after processing
                {
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    string prefix = ReaderUtils.DiscoverPrefix(xr, "theme");
                    XmlNodeList themes = ReaderUtils.GetElementsByTagName(xr, "theme", prefix);
                    string themeName = ReaderUtils.GetAttribute(themes[0], "name"); // If this fails, something is completely wrong
                    Workbook.WorkbookTheme = new Theme(themeName);
                    ColorScheme colorScheme = new ColorScheme();
                    Workbook.WorkbookTheme.Colors = colorScheme;
                    XmlNodeList colors = ReaderUtils.GetElementsByTagName(xr, "clrScheme", prefix);

                    foreach (XmlNode color in colors)
                    {
                        string colorSchemeName = ReaderUtils.GetAttribute(color, "name", "");
                        Workbook.WorkbookTheme.Colors.Name = colorSchemeName;
                        XmlNodeList colorNodes = color.ChildNodes;
                        foreach (XmlNode colorNode in colorNodes)
                        {
                            string name = colorNode.LocalName;
                            switch (name)
                            {
                                case "dk1":
                                    colorScheme.Dark1 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "lt1":
                                    colorScheme.Light1 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "dk2":
                                    colorScheme.Dark2 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "lt2":
                                    colorScheme.Light2 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "accent1":
                                    colorScheme.Accent1 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "accent2":
                                    colorScheme.Accent2 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "accent3":
                                    colorScheme.Accent3 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "accent4":
                                    colorScheme.Accent4 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "accent5":
                                    colorScheme.Accent5 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "accent6":
                                    colorScheme.Accent6 = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "hlink":
                                    colorScheme.Hyperlink = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "folHlink":
                                    colorScheme.FollowedHyperlink = ParseColor(colorNode.ChildNodes);
                                    break;
                            }

                        }
                    }
                    RederPlugInHandler.HandleInlineQueuePlugins(ref stream, Workbook, PlugInUUID.ThemeInlineReader);
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Parses a color value (either RGB-like or enumerated system color)
        /// </summary>
        /// <param name="childNodes">List of XML nodes that can contain color values</param>
        /// <returns><see cref="IColor"/> value or null, if no color could be determined</returns>
        private IColor ParseColor(XmlNodeList childNodes)
        {
            foreach (XmlNode node in childNodes)
            {
                if (node.LocalName == "sysClr")
                {
                    SystemColor.Value value = ParseSystemColor(node);
                    SystemColor systemColor = new SystemColor();
                    systemColor.ColorValue = value;
                    string lastColor = ReaderUtils.GetAttribute(node, "lastClr");
                    if (lastColor != null)
                    {
                        systemColor.LastColor = lastColor;
                    }
                    return systemColor;
                }
                else if (node.LocalName == "srgbClr")
                {
                    SrgbColor srgbColor = new SrgbColor();
                    srgbColor.ColorValue = ReaderUtils.GetAttribute(node, "val");
                    return srgbColor;
                }
            }
            return null;
        }

        /// <summary>
        /// Tries to parse a system color
        /// </summary>
        /// <param name="innerNode">Color scheme sub-node</param>
        /// <returns>System color</returns>
        /// <exception cref="NanoXLSX.Exceptions.StyleException">Throws IOException in case of an invalid value</exception>
        private static SystemColor.Value ParseSystemColor(XmlNode innerNode)
        {
            string value = ReaderUtils.GetAttribute(innerNode, "val");
            if (string.IsNullOrEmpty(value))
            {
                throw new IOException("The system color entry was null or empty");
            }
            try
            {
                return SystemColor.MapStringToValue(value);
            }
            catch (Exception ex)
            {
                throw new IOException("The system color entry '" + value + "' could not be parsed", ex);
            }
        }
        #endregion
    }
}
