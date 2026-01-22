/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Xml;
using static NanoXLSX.Enums.Password;



namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by password readers
    /// </summary>
    public interface IPasswordReader : IPassword
    {

        /// <summary>
        /// Method to initialize the password reader
        /// </summary>
        /// <param name="type">Target type of the password writer</param>
        /// <param name="readerOptions">Reader options</param>
        void Init(PasswordType type, ReaderOptions readerOptions);

        /// <summary>
        /// Reads the attributes of the passed XML node that contains password information
        /// </summary>
        /// <param name="node">XML node</param>
        void ReadXmlAttributes(XmlNode node);
    }
}
