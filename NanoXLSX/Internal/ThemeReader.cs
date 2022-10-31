using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Themes;
using IOException = NanoXLSX.Shared.Exceptions.IOException;

namespace NanoXLSX.Internal
{
    public class ThemeReader
    {
        public Theme CurrentTheme { get; set; }

  
        #region methods

        /// <summary>
        /// Reads the XML file form the passed stream and processes the theme file (if available)
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <param name="number">NUmber of the theme. Default is 1</param>
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an error</exception>
        public void Read(MemoryStream stream, int number)
        {
            try
            {
                using (stream) // Close after processing
                {
                    CurrentTheme = new Theme(number);
                    ColorScheme colorScheme = new ColorScheme(number);
                    CurrentTheme.Colors = colorScheme;
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    XmlNodeList colors = xr.GetElementsByTagName("clrScheme");
                    foreach (XmlNode color in colors)
                    {
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
                                    colorScheme.HyperLink = ParseColor(colorNode.ChildNodes);
                                    break;
                                case "folHlink":
                                    colorScheme.FollowedHyperlink = ParseColor(colorNode.ChildNodes);
                                    break;
                            }
                            
                        }
                        
                        /*
                        string rowAttribute = ReaderUtils.GetAttribute(row, "r");
                        if (rowAttribute != null)
                        {
                            string nameAttribute = ReaderUtils.GetAttribute(row, "name");
                            RowDefinition.AddRowDefinition(Rows, rowAttribute, null, hiddenAttribute);
                            string heightAttribute = ReaderUtils.GetAttribute(row, "ht");
                            RowDefinition.AddRowDefinition(Rows, rowAttribute, heightAttribute, null);
                        }
                        if (row.HasChildNodes)
                        {
                            foreach (XmlNode rowChild in row.ChildNodes)
                            {
                                ReadCell(rowChild);
                            }
                        }
                    }
                    GetSheetView(xr);
                    GetMergedCells(xr);
                    GetSheetFormats(xr);
                    GetAutoFilters(xr);
                    GetColumns(xr);
                    GetSheetProtection(xr);
                        
                }
                    */
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        private IColor ParseColor(XmlNodeList childNodes)
        {
            foreach(XmlNode node in childNodes)
            {
                if (node.LocalName == "sysClr")
                {
                    SystemColor.Value value = ParseSystemColor(node);
                    SystemColor systemColor = new SystemColor();
                    systemColor.ColorValue = value;
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
        /// <exception cref="NanoXLSX.Shared.Exceptions.IOException">Throws IOException in case of an invalid value</exception>
        private static SystemColor.Value ParseSystemColor(XmlNode innerNode)
        {
            string value = ReaderUtils.GetAttribute(innerNode, "val");
            if (value != null)
            {
                SystemColor.Value systemColor;
                if (Enum.TryParse(value, out systemColor))
                {
                    return systemColor;
                }
            }
            throw new IOException("The system color entry'" + value + "' is not valid");
        }

        #endregion

    }
}
