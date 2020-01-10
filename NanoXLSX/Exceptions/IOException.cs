/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Exceptions
{
    /// <summary>
    /// Class for exceptions regarding stream or save error incidents
    /// </summary>
    [Serializable]
    public class IOException : Exception
    {
        /// <summary>
        /// Gets or sets the title of the exception
        /// </summary>
        public string ExceptionTitle { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public IOException()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="title">Title of the exception</param>
        public IOException(string title, string message)
            : base(title + ": " + message)
        { ExceptionTitle = title; }
        /// <summary>
        /// Constructor with passed message and inner exception
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="inner">Inner exception</param>
        /// <param name="title">Title of the exception</param>
        public IOException(string title, string message, Exception inner)
            : base(message, inner)
        { ExceptionTitle = title; }
    }
}