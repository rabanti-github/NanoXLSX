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
    /// Class for exceptions regarding range incidents (e.g. out-of-range)
    /// </summary>
    [Serializable]
    public class RangeException : System.Exception
    {
        /// <summary>
        /// Gets or sets the title of the exception
        /// </summary>
        public string ExceptionTitle { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public RangeException() : base()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        /// <param name="title">Title of the exception</param>
        public RangeException(string title, string message)
            : base(title + ": " + message)
        { this.ExceptionTitle = title; }
    }
}