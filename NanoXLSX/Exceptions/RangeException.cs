/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2021
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace NanoXLSX.Exceptions
{
    /// <summary>
    /// Class for exceptions regarding range incidents (e.g. out-of-range)
    /// </summary>
    [Serializable]
    public class RangeException : Exception
    {
        /// <summary>
        /// Default constructor
        /// </summary>
        public RangeException()
        { }
        /// <summary>
        /// Constructor with passed message
        /// </summary>
        /// <param name="message">Message of the exception</param>
        public RangeException(string message)
            : base( message)
        {  }
    }
}
