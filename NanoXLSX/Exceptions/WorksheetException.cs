/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Runtime.CompilerServices;

namespace NanoXLSX.Exceptions
{
    /// <summary>
    /// Class for exceptions regarding worksheet incidents
    /// </summary>
    [Serializable]
    public class WorksheetException : Exception
    {
        /// <summary>
        /// Gets or sets the title of the exception
        /// </summary>
        public string ExceptionTitle { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public WorksheetException()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="title">Title of the exception</param>
        public WorksheetException(string title, string message)
            : base(title + ": " + message)
        { ExceptionTitle = title; }
    }

    #region doc
    /// <summary>
    /// Sub-namespace for exceptions
    /// </summary>
    [CompilerGenerated]
    class NamespaceDoc // This class is only for documentation purpose (Sandcastle)
    { }
    #endregion

}