using System;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Themes;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for theme definitions of XLSX files.
    /// </summary>
    public class ThemeReader : IPlugInReader
    {

        /// <summary>
        /// Currently active theme
        /// </summary>
        public Theme CurrentTheme { get; set; }


        #region methods

        /// <summary>
        /// Reads the XML file form the passed stream and processes the theme file (if available)
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <exception cref="NanoXLSX.Exceptions.IOException">Throws IOException in case of an error</exception>
        public virtual void Read(MemoryStream stream)
        {
            PreRead(stream);
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
                    CurrentTheme = new Theme(themeName);
                    ColorScheme colorScheme = new ColorScheme();
                    CurrentTheme.Colors = colorScheme;
                    XmlNodeList colors = ReaderUtils.GetElementsByTagName(xr, "clrScheme", prefix);
                    
                    foreach (XmlNode color in colors)
                    {
                        string colorSchemeName = ReaderUtils.GetAttribute(color, "name", "");
                        CurrentTheme.Colors.Name = colorSchemeName;
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
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
            PostRead(stream);
        }

        /// <summary>
        /// Method that is called before the <see cref="Read(MemoryStream)"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="stream">Stream of the XML file. The stream must be reset in this method at the end, if any stream opeartion was performed</param>
        public virtual void PreRead(MemoryStream stream)
        {
            // NoOp - replaced by plugIn
        }

        /// <summary>
        /// Method that is called after the <see cref="Read(MemoryStream)"/> method is executed. 
        /// This virtual method is empty by default and can be overridden by a plug-in package
        /// </summary>
        /// <param name="stream">Stream of the XML file. The stream must be reset in this method before any stream operation is performed</param>
        public virtual void PostRead(MemoryStream stream)
        {
            // NoOp - replaced by plugIn
        }

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
