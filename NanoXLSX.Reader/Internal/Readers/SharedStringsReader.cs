/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using NanoXLSX.Utils.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for the shared strings table of XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.SharedStringsReader)]
    public class SharedStringsReader : ISharedStringReader
    {

        #region privateFields
        private bool capturePhoneticCharacters;
        private readonly List<PhoneticInfo> phoneticsInfo;
        private Stream stream;
        #endregion

        #region properties

        /// <summary>
        /// List of shared string entries
        /// </summary>
        /// <value>
        /// String entry, sorted by its internal index of the table
        /// </value>
        public List<string> SharedStrings { get; private set; }

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
        public SharedStringsReader()
        {
            phoneticsInfo = new List<PhoneticInfo>();
            SharedStrings = new List<string>();
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
            if (readerOptions is ReaderOptions options)
            {
                this.capturePhoneticCharacters = options.EnforcePhoneticCharacterImport;
            }
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
                    StringBuilder sb = new StringBuilder();
                    using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                    {
                        while (reader.Read())
                        {
                            if (!XmlStreamUtils.IsElement(reader, "si"))
                            {
                                continue;
                            }
                            sb.Clear();
                            ReadSiElement(reader, sb);
                            if (capturePhoneticCharacters)
                            {
                                SharedStrings.Add(ProcessPhoneticCharacters(sb));
                            }
                            else
                            {
                                SharedStrings.Add(sb.ToString());
                            }
                        }
                        InlinePluginHandler?.Invoke(stream, Workbook, PlugInUUID.SharedStringsInlineReader, Options, null);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Reads a single shared-string item (&lt;si&gt;) using a subtree reader.
        /// Collects text from all &lt;t&gt; elements, skipping or capturing &lt;rPh&gt; phonetic elements.
        /// </summary>
        private void ReadSiElement(XmlReader reader, StringBuilder sb)
        {
            using (XmlReader siSubtree = reader.ReadSubtree())
            {
                siSubtree.Read(); // consume the <si> open tag
                while (siSubtree.Read())
                {
                    if (siSubtree.NodeType != XmlNodeType.Element)
                    {
                        continue;
                    }
                    if (siSubtree.LocalName.Equals("rPh", StringComparison.OrdinalIgnoreCase))
                    {
                        if (capturePhoneticCharacters)
                        {
                            ReadPhoneticElement(siSubtree);
                        }
                        else
                        {
                            using (siSubtree.ReadSubtree()) { } // dispose immediately; positions siSubtree at </rPh>
                        }
                    }
                    else if (siSubtree.LocalName.Equals("t", StringComparison.OrdinalIgnoreCase))
                    {
                        string text;
                        using (XmlReader tSubtree = siSubtree.ReadSubtree())
                        {
                            tSubtree.Read(); // position at <t>
                            text = tSubtree.ReadElementContentAsString();
                        }
                        if (!string.IsNullOrEmpty(text))
                        {
                            sb.Append(text);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Reads a &lt;rPh&gt; element and captures its text and position attributes for phonetic annotation.
        /// Uses a subtree reader so the outer reader advances past the element on return.
        /// </summary>
        private void ReadPhoneticElement(XmlReader reader)
        {
            string start = reader.GetAttribute("sb");
            string end = reader.GetAttribute("eb");
            string text = null;
            using (XmlReader rPhSubtree = reader.ReadSubtree())
            {
                rPhSubtree.Read(); // consume the <rPh> open tag
                while (rPhSubtree.Read())
                {
                    if (rPhSubtree.NodeType == XmlNodeType.Element
                        && rPhSubtree.LocalName.Equals("t", StringComparison.OrdinalIgnoreCase))
                    {
                        text = rPhSubtree.ReadElementContentAsString();
                    }
                }
            }
            if (!string.IsNullOrEmpty(text))
            {
                phoneticsInfo.Add(new PhoneticInfo(text, start, end));
            }
        }

        /// <summary>
        /// Function to add determined phonetic tokens
        /// </summary>
        /// <param name="sb">Original StringBuilder</param>
        /// <returns>Text with added phonetic characters (after particular characters, in brackets)</returns>
        private string ProcessPhoneticCharacters(StringBuilder sb)
        {
            if (phoneticsInfo.Count == 0)
            {
                return sb.ToString();
            }
            string text = sb.ToString();
            StringBuilder sb2 = new StringBuilder();
            int currentTextIndex = 0;
            foreach (PhoneticInfo info in phoneticsInfo)
            {
                sb2.Append(text.Substring(currentTextIndex, info.StartIndex + info.Length - currentTextIndex));
                sb2.Append('(').Append(info.Value).Append(')');
                currentTextIndex = info.StartIndex + info.Length;
            }
            sb2.Append(text.Substring(currentTextIndex));

            phoneticsInfo.Clear();
            return sb2.ToString();
        }


        #endregion

        #region sub-classes
        /// <summary>
        /// Class to represent a phonetic transcription of character sequence.
        /// Note: Invalid values will lead to a crash. The specifications requires a start index, an end index and a value
        /// </summary>
        sealed class PhoneticInfo
        {
            /// <summary>
            /// Transcription value
            /// </summary>
            public string Value { get; private set; }
            /// <summary>
            /// Absolute start index within the original string
            /// </summary>
            public int StartIndex { get; private set; }
            /// <summary>
            /// Number of characters of the original string that are described by this transcription token
            /// </summary>
            public int Length { get; private set; }

            /// <summary>
            /// Constructor with parameters
            /// </summary>
            /// <param name="value">Transcription value</param>
            /// <param name="start">Absolute start index as string</param>
            /// <param name="end">Absolute end index as string</param>
            public PhoneticInfo(string value, string start, string end)
            {
                Value = value;
                StartIndex = ParserUtils.ParseInt(start);
                Length = ParserUtils.ParseInt(end) - StartIndex;

            }
        }
        #endregion
    }
}
