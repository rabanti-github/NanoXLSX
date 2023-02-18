/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
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

        public List<Relationship> Relationships { get; set; } = new List<Relationship>();

        #endregion


        #region functions

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
                    var id = ReaderUtils.GetAttribute(relationship, "Id");
                    var type = ReaderUtils.GetAttribute(relationship, "Type");
                    var target = ReaderUtils.GetAttribute(relationship, "Target");
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

        public class Relationship
        {
            public string Id { get; set; }
            public string Type { get; set; }
            public string Target { get; set; }
        }
    }
}