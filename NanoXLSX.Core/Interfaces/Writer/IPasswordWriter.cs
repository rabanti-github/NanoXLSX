/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Utils.Xml;
using System.Collections.Generic;
using static NanoXLSX.Enums.Password;


namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by specific writers that provides password handling
    /// </summary>
    public interface IPasswordWriter : IPassword
    {

        /// <summary>
        /// Gets the target type of the password
        /// </summary>
        PasswordType Type { get; }

        /// <summary>
        /// Method to initialize the password writer
        /// </summary>
        /// <param name="type">Target type of the password writer</param>
        /// <param name="passwordHash">Hash that will be written</param>
        void Init(PasswordType type, string passwordHash);

        /// <summary>
        /// Gets an IEnumerable of XML attributes
        /// </summary>
        /// <returns></returns>
        IEnumerable<XmlAttribute> GetAttributes();
    }
}
