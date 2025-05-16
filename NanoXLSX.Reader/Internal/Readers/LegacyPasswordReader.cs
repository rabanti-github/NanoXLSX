/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Xml;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using static NanoXLSX.Internal.Enums.ReaderPassword;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Class representing a reader for legacy passwords
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.PASSWORD_READER)]
    public class LegacyPasswordReader : IPasswordReader
    {
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
        /// Reader options
        /// </summary>
        public ReaderOptions Options { get; private set; }

        /// <summary>
        /// Gets or sets the password hash
        /// </summary>
        public string PasswordHash
        {
            get { return passwordHash; }
            set { passwordHash = value; }
        }


        /// <summary>
        /// Default constructor - Must be defined for instantiation of the plug-ins
        /// </summary>
        internal LegacyPasswordReader()
        {
        }

        /// <summary>
        /// Initialization method (interface implementation)
        /// </summary>
        /// <param name="type">Password type</param>
        /// <param name="options">Reader options</param> 
        public void Init(PasswordType type, ReaderOptions options)
        {
            this.Type = type;
            this.Options = options;
        }

        /// <summary>
        /// Reads the attributes of the passed XML node that contains password information
        /// </summary>
        /// <param name="node">XML node</param>
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
                if (Options.IgnoreNotSupportedPasswordAlgorithms)
                {
                    this.ContemporaryAlgorithmDetected = true;
                }
                else
                {
                    throw new NotSupportedContentException("A not supported, contemporary password algorithm for the worksheet protection was detected. Check possible packages to add support to NanoXLSX, or ignore this error by a reader option");
                }
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
        /// Gets the password. This method is not supported in a reader and will always return null
        /// </summary>
        /// <returns>Always null, since the plain text password cannot be recovered</returns>
        public string GetPassword()
        {
            return null; // The reader cannot recover the plain text password
        }

        /// <summary>
        /// Indicates whether a password is set. This can be the case, if a legacy or contemporary password was detected, regardless of the ability of the decoding of this reader
        /// </summary>
        /// <returns>True if a password was set</returns>
        public bool PasswordIsSet()
        {
            return passwordHash != null || ContemporaryAlgorithmDetected;
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
