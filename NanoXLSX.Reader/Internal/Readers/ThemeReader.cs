/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Xml;
using NanoXLSX.Colors;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Themes;
using NanoXLSX.Utils.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for theme definitions of XLSX files.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.ThemeReader)]
    public class ThemeReader : IPluginBaseReader
    {

        private Stream stream;

        #region properties
        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get; set; }
        /// <summary>
        /// Reader options
        /// </summary>
        public IOptions Options { get; set; }
        /// <summary>
        /// Reference to the <see cref="ReaderPlugInHandler"/>, to be used for post operations in the <see cref="Execute"/> method
        /// </summary>
        public Action<Stream, Workbook, string, IOptions, int?> InlinePluginHandler { get; set; }
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
        /// <param name="inlinePluginHandler">Reference to the a handler action, to be used for post operations in reader methods</param>
        public void Init(Stream stream, Workbook workbook, IOptions readerOptions, Action<Stream, Workbook, string, IOptions, int?> inlinePluginHandler)
        {
            this.stream = stream;
            this.Workbook = workbook;
            this.Options = readerOptions;
            this.InlinePluginHandler = inlinePluginHandler;
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
                    ColorScheme colorScheme = null;
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType != XmlNodeType.Element)
                            {
                                continue;
                            }
                            if (XmlStreamUtils.IsElement(reader, "theme") && colorScheme == null)
                            {
                                string themeName = reader.GetAttribute("name");
                                Workbook.WorkbookTheme = new Theme(themeName);
                                colorScheme = new ColorScheme();
                                Workbook.WorkbookTheme.Colors = colorScheme;
                            }
                            else if (XmlStreamUtils.IsElement(reader, "clrScheme") && colorScheme != null)
                            {
                                ReadClrScheme(reader, colorScheme);
                            }
                        }
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.ThemeInlineReader, Options, null);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Reads the clrScheme element: captures its name attribute and iterates its color-slot
        /// children (dk1, lt1, …) using a single subtree pass to avoid sibling-skipping.
        /// </summary>
        private static void ReadClrScheme(XmlReader reader, ColorScheme colorScheme)
        {
            colorScheme.Name = reader.GetAttribute("name") ?? string.Empty;
            using (XmlReader subtree = reader.ReadSubtree())
            {
                subtree.Read(); // consume the <clrScheme> open tag
                while (subtree.Read())
                {
                    if (subtree.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }
                    string slot = subtree.LocalName;
                    IColor color = ReadColorEntry(subtree);
                    switch (slot)
                    {
                        case "dk1":
                            colorScheme.Dark1 = color;
                            break;
                        case "lt1":
                            colorScheme.Light1 = color;
                            break;
                        case "dk2":
                            colorScheme.Dark2 = color;
                            break;
                        case "lt2":
                            colorScheme.Light2 = color;
                            break;
                        case "accent1":
                            colorScheme.Accent1 = color;
                            break;
                        case "accent2":
                            colorScheme.Accent2 = color;
                            break;
                        case "accent3":
                            colorScheme.Accent3 = color;
                            break;
                        case "accent4":
                            colorScheme.Accent4 = color;
                            break;
                        case "accent5":
                            colorScheme.Accent5 = color;
                            break;
                        case "accent6":
                            colorScheme.Accent6 = color;
                            break;
                        case "hlink":
                            colorScheme.Hyperlink = color;
                            break;
                        case "folHlink":
                            colorScheme.FollowedHyperlink = color;
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Reads the color value from the children of a color-slot element (dk1, lt1, etc.).
        /// Uses depth-based traversal so the reader is left on the slot's end element on return,
        /// allowing the caller's while loop to correctly advance to the next slot sibling.
        /// </summary>
        private static IColor ReadColorEntry(XmlReader reader)
        {
            if (reader.IsEmptyElement)
            {
                return null;
            }
            int slotDepth = reader.Depth;
            while (reader.Read())
            {
                if (reader.Depth == slotDepth)
                {
                    break; // at the slot's end element — caller's Read() moves to next sibling
                }
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.LocalName.Equals("sysClr", StringComparison.OrdinalIgnoreCase))
                    {
                        string val = reader.GetAttribute("val");
                        SystemColor systemColor = new SystemColor
                        {
                            ColorValue = ParseSystemColor(val)
                        };
                        string lastColor = reader.GetAttribute("lastClr");
                        if (lastColor != null)
                        {
                            systemColor.LastColor = lastColor;
                        }
                        reader.Skip(); // advance past sysClr to the slot's end element
                        return systemColor;
                    }
                    else if (reader.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase))
                    {
                        SrgbColor color = new SrgbColor
                        {
                            ColorValue = reader.GetAttribute("val")
                        };
                        reader.Skip(); // advance past srgbClr to the slot's end element
                        return color;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Tries to parse a system color string value
        /// </summary>
        /// <param name="value">String value of the val attribute</param>
        /// <returns>System color enum value</returns>
        /// <exception cref="NanoXLSX.Exceptions.IOException">Throws IOException in case of an invalid value</exception>
        private static SystemColor.Value ParseSystemColor(string value)
        {
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
