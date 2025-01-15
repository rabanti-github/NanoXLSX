using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;

namespace NanoXLSX.Internal.Readers
{
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
            return passwordHash != null;
        }

        public void ReadXmlAttributes(XmlNode node)
        {
            string attribute = ReaderUtils.GetAttribute(node, "algorithmName");
            if (attribute != null)
            {
                this.ContemporaryAlgorithmDetected = true;
            }

            if (Type == PasswordType.WORKBOOK_PROTECTION)
            {
                attribute = ReaderUtils.GetAttribute(node, "workbookPassword");
                if (attribute != null)
                {
                    this.passwordHash = attribute;
                }
            }
            else
            {
                attribute = ReaderUtils.GetAttribute(node, "password");
                if (attribute != null)
                {
                    this.passwordHash = attribute;
                }
            }
        }

        public void SetPassword(string plainText)
        {
            throw new NotImplementedException();
        }

        public void UnsetPassword()
        {
            throw new NotImplementedException();
        }
    }
}
