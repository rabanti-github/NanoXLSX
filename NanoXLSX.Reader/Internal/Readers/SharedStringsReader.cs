/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using NanoXLSX.Interfaces.Plugin;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for the shared strings table of XLSX files
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.SHARED_STRINGS_READER)]
    public class SharedStringsReader : ISharedStringReader
    {

        #region privateFields
        private bool capturePhoneticCharacters = false;
        private readonly List<PhoneticInfo> phoneticsInfo = null;
        private MemoryStream stream;
        private Workbook workbook;
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
        /// Gets whether the workbook contains shared strings
        /// </summary>
        /// <value>
        /// True if at least one shared string object exists in the workbook
        /// </value>
        public bool HasElements
        {
            get
            {
                return SharedStrings.Count > 0;
            }
        }

        /// <summary>
        /// Workbook reference where read data is stored (should not be null)
        /// </summary>
        public Workbook Workbook { get => workbook; set => workbook = value; }

        /// <summary>
        /// Gets the value of the shared string table by its index
        /// </summary>
        /// <param name="index">Index of the stared string entry</param>
        /// <returns>Determined shared string value. Returns null in case of a invalid index</returns>
        public string GetString(int index)
        {
            if (!HasElements || index > SharedStrings.Count - 1 || index < 0)
            {
                return null;
            }
            return SharedStrings[index];
        }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        public SharedStringsReader()
        {
            phoneticsInfo = new List<PhoneticInfo>();
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
            this.workbook = workbook;
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
                    XmlDocument xr = new XmlDocument();
                    xr.XmlResolver = null;
                    xr.Load(stream);
                    StringBuilder sb = new StringBuilder();
                    foreach (XmlNode node in xr.DocumentElement.ChildNodes)
                    {
                        if (node.LocalName.Equals("si", StringComparison.InvariantCultureIgnoreCase))
                        {
                            sb.Clear();
                            GetTextToken(node, ref sb);
                            if (capturePhoneticCharacters)
                            {
                                SharedStrings.Add(ProcessPhoneticCharacters(sb));
                            }
                            else
                            {
                                SharedStrings.Add(sb.ToString());
                            }
                        }
                    }
                    RederPlugInHandler.HandleInlineQueuePlugins(ref stream, Workbook, PlugInUUID.SHARED_STRINGS_INLINE_READER);
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the " + nameof(stream) + ". Please see the inner exception:", ex);
            }
        }

        /// <summary>
        /// Function collects text tokens recursively in case of a split by formatting
        /// </summary>
        /// <param name="node">Root node to process</param>
        /// <param name="sb">StringBuilder reference</param>
        private void GetTextToken(XmlNode node, ref StringBuilder sb)
        {
            if (node.LocalName.Equals("rPh", StringComparison.InvariantCultureIgnoreCase))
            {
                if (capturePhoneticCharacters && !string.IsNullOrEmpty(node.InnerText))
                {
                    string start = node.Attributes.GetNamedItem("sb").InnerText;
                    string end = node.Attributes.GetNamedItem("eb").InnerText;
                    phoneticsInfo.Add(new PhoneticInfo(node.InnerText, start, end));
                }
                return;
            }

            if (node.LocalName.Equals("t", StringComparison.InvariantCultureIgnoreCase) && !string.IsNullOrEmpty(node.InnerText))
            {
                sb.Append(node.InnerText);
            }
            if (node.HasChildNodes)
            {
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    GetTextToken(childNode, ref sb);
                }
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
                sb2.Append("(").Append(info.Value).Append(")");
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
