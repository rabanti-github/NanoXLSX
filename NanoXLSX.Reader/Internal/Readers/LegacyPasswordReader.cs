/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;

namespace NanoXLSX.Internal.Readers
{
     /// <summary>
     /// Class representing a reader for legacy passwords
     /// </summary>
    public class LegacyPasswordReader : IPasswordReader
    {
        /// <summary>
        /// Target type of the password
        /// </summary>
        public enum PasswordType
        {
            /// <summary>
            /// Password is used to protect a workbook
            /// </summary>
            WORKBOOK_PROTECTION,
            /// <summary>
            /// Password is used to protect a worksheet
            /// </summary>
            WORKSHEET_PROTECTION
        }

        private string passwordHash = null;

        /// <summary>
        /// Gets whether a contemporary password algorithm was detected (not supported by core functionality)
        /// </summary>
        public bool ContemporaryAlgorithmDetected { get; private set; }

        /// <summary>
        /// Current target type of the password instance
        /// </summary>
        public PasswordType Type { get; private set; }

        /// <summary>
        /// Gets or sets the password hash
        /// </summary>
        public string PasswordHash 
        { 
            get { return passwordHash; } 
            set { passwordHash = value; } 
        }

        /// <summary>
        /// Default constructor with parameter
        /// </summary>
        /// <param name="type">Current target type of the password instance</param>
        public LegacyPasswordReader(PasswordType type)
        {
            this.Type = type;
        }

        /// <summary>
        /// Not relevant for the reader (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <param name="passwordInstance">Source instance</param>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void CopyFrom(IPassword passwordInstance)
        {
            throw new NotImplementedException();
        }

        public string GetPassword()
        {
            return null; // The reader cannot recover the plain text password
        }

        public bool PasswordIsSet()
        {
            return passwordHash != null || ContemporaryAlgorithmDetected;
        }

        public void ReadXmlAttributes(XmlNode node)
        {
            string attribute = null;
            if (Type == PasswordType.WORKBOOK_PROTECTION)
            {
                attribute = ReaderUtils.GetAttribute(node, "workbookAlgorithmName");
            }
            else
            {
                attribute = ReaderUtils.GetAttribute(node, "algorithmName");
            }
            if (attribute != null)
            {
                this.ContemporaryAlgorithmDetected = true;
            }

            if (Type == PasswordType.WORKBOOK_PROTECTION)
            {
                attribute = ReaderUtils.GetAttribute(node, "workbookPassword");
                if (attribute != null)
                {
                    this.PasswordHash = attribute;
                }
            }
            else
            {
                attribute = ReaderUtils.GetAttribute(node, "password");
                if (attribute != null)
                {
                    this.PasswordHash = attribute;
                }
            }
        }

        /// <summary>
        /// Not relevant for the reader (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <param name="plainText"></param>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void SetPassword(string plainText)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Not relevant for the reader (inherited from <see cref="IPassword"/>)
        /// </summary>
        /// <exception cref="NotImplementedException">Throws a NotImplementedException if called in any case</exception>
        public void UnsetPassword()
        {
            throw new NotImplementedException();
        }
    }
}
