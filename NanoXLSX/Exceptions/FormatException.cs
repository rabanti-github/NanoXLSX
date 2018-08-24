/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
using System;

namespace NanoXLSX.Exception
{
    /// <summary>
    /// Class for exceptions regarding format error incidents
    /// </summary>
    [Serializable]
    public class FormatException : System.Exception
    {
        /// <summary>
        /// Gets or sets the title of the exception
        /// </summary>
        public string ExceptionTitle { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public FormatException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public FormatException(string message)
            : base(message)
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="title">Title of the exception</param>
        public FormatException(string title, string message)
            : base(title + ": " + message)
        { this.ExceptionTitle = title; }
        /// <summary>
        /// Constructor with passed message and inner exception
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="inner">Inner exception</param>
        /// <param name="title">Title of the exception</param>
        public FormatException(string title, string message, System.Exception inner)
            : base(message, inner)
        { this.ExceptionTitle = title; }
    }
}