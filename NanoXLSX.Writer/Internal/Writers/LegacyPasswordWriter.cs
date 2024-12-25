using System;
using System.Collections.Generic;
using System.Text;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Writer;

namespace NanoXLSX.Internal.Writers
{
    public class LegacyPasswordWriter : IPasswordWriter
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

        private string passwordHash;

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
        public LegacyPasswordWriter(PasswordType type, string hash)
        {
            this.Type = type;
            this.passwordHash = hash;
        }

        public void CopyFrom(IPassword passwordInstance)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Gets the internally used XML attributes that are set when a workbook is saved
        /// </summary>
        /// <returns>XML attributes with a leading space, or an empty string, if no password was set</returns>
        /// \remark <remarks>This method is only internally used on writing workbooks</remarks>
        public new string GetXmlAttributes()
        {
            if (PasswordIsSet())
            {
                if (Type == PasswordType.WORKSHEET_PROTECTION)
                {
                    return " password =\"" + passwordHash + "\"";
                }
                else
                {
                    return " workbookPassword=\"" + passwordHash + "\"";
                }
            }
            return "";
        }

        public bool PasswordIsSet()
        {
            return passwordHash != null;
        }

        public void SetPassword(string plainText)
        {
            throw new NotImplementedException();
        }

        public void UnsetPassword()
        {
            throw new NotImplementedException();
        }

        public string GetPassword()
        {
            throw new NotImplementedException();
        }
    }
}
