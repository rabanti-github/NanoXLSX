/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.LowLevel
{
    /// <summary>
    /// Class representing a reader for relationship of XLSX files
    /// </summary>
    public class RelationshipReader
    {
        #region properties

        /// <summary>
        /// List of workbook relationship entries
        /// </summary>
        public List<Relationship> Relationships { get; set; } = new List<Relationship>();

        #endregion


        #region functions

        /// <summary>
        /// Reads the XML file form the passed stream and processes the workbook relationship document
        /// </summary>
        /// <param name="stream">Stream of the XML file</param>
        /// <exception cref="Exceptions.IOException">Throws IOException in case of an error</exception>
        public void Read(MemoryStream stream)
        {
            if (stream == null) return;
            try
            {
                XmlDocument xr;
                using (stream) // Close after processing
                {
                    xr = new XmlDocument
                    {
                        XmlResolver = null
                    };
                    xr.Load(stream);
                }

                var relationships = xr.GetElementsByTagName("Relationship");
                foreach (XmlNode relationship in relationships)
                {
                    string id = ReaderUtils.GetAttribute(relationship, "Id");
                    string type = ReaderUtils.GetAttribute(relationship, "Type");
                    string target = ReaderUtils.GetAttribute(relationship, "Target");
                    if (target.StartsWith("/"))
                    {
                        target = target.TrimStart('/');
                    }
                    if (!target.StartsWith("xl/"))
                    {
                        target = "xl/" + target;
                    }
                    Relationships.Add(
                        new Relationship
                        {
                            Id = id,
                            Type = type,
                            Target = target,
                        });
                }
            }
            catch (Exception ex)
            {
                throw new IOException("The XML entry could not be read from the input stream. Please see the inner exception:", ex);
            }
        }

        #endregion

        #region sub-classes
        /// <summary>
        /// Class to represent a workbook relation
        /// </summary>
        public class Relationship
        {
            /// <summary>
            /// UD of the relation
            /// </summary>
            public string Id { get; set; }
            /// <summary>
            /// Type of the relation
            /// </summary>
            public string Type { get; set; }
            /// <summary>
            /// Target of the relation
            /// </summary>
            public string Target { get; set; }
        }
        #endregion
    }
}