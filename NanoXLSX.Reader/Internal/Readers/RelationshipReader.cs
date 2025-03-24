﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using NanoXLSX.Interfaces.Reader;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{

    /// <summary>
    /// Class representing a reader for relationship of XLSX files
    /// </summary>
    public class RelationshipReader : IPlugInReader
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
        /// \remark <remarks>This method is virtual. Plug-in packages may override it</remarks>
        /// <exception cref="IOException">Throws IOException in case of an error</exception>
        public virtual void Read(MemoryStream stream)
        {
            PreRead(stream);
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
