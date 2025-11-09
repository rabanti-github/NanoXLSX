/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;
using static NanoXLSX.Internal.Enums.WriterPassword;

namespace NanoXLSX.Internal.Writers
{
    /// <summary>
    /// Class to write a legacy password
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.PasswordWriter)]
    public class LegacyPasswordWriter : IPasswordWriter
    {
        #region properties

        /// <summary>
        /// Current target type of the password instance
        /// </summary>
        public PasswordType Type { get; private set; }

        /// <summary>
        /// Gets or sets the password hash
        /// </summary>
        public string PasswordHash { get; set; }

        /// <summary>
        /// Default constructor with parameter
        /// </summary>
        /// <param name="type">Target type of the password instance</param>
        /// <param name="hash">Hash representation of the password (do not use null)</param>
        public LegacyPasswordWriter(PasswordType type, string hash)
        {
            this.Type = type;
            this.PasswordHash = hash;
        }

        #endregion
        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public LegacyPasswordWriter()
        {
        }

        #endregion
        #region methods
        /// <summary>
        /// Initializer method with all mandatory parameters
        /// </summary>
        /// <param name="type">Target type of the password instance</param>
        /// <param name="passwordHash">Hash representation of the password (do not use null)</param>
        public void Init(PasswordType type, string passwordHash)
        {
            this.Type = type;
            this.PasswordHash = passwordHash;
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <param name="passwordInstance">Source instance</param>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void CopyFrom(IPassword passwordInstance)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets whether a password to write is defined
        /// </summary>
        /// <returns>True if a password is set to be written</returns>
        public bool PasswordIsSet()
        {
            return PasswordHash != null;
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public string GetPassword()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <param name="plainText"></param>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void SetPassword(string plainText)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Not relevant for the writer (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void UnsetPassword()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the XML attributes of the current password instance, that are used when writing XLSX files
        /// </summary>
        /// <returns>IENumerable of attributes</returns>
        public IEnumerable<XmlAttribute> GetAttributes()
        {
            List<XmlAttribute> attributes = new List<XmlAttribute>();
            if (Type == PasswordType.WORKSHEET_PROTECTION)
            {
                attributes.Add(XmlAttribute.CreateAttribute("password", PasswordHash));
            }
            else
            {
                attributes.Add(XmlAttribute.CreateAttribute("workbookPassword", PasswordHash));
            }
            return attributes;
        }
        #endregion
    }
}
