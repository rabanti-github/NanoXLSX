/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Text;

namespace NanoXLSX.Shared.Interfaces
{
    /// <summary>
    /// Interface to represent complex text data that can be formatted somehow
    /// </summary>
    public interface IFormattableText
    {

        /// <summary>
        /// Method to add the formatted text value, from the object of the class that has implemented the interface, to the passed string builder
        /// </summary>
        /// <param name="sb">String builder instance</param>
        /// <remarks>The formatted value must be XLM-attribute-ready, means already escaped and containing the right 
        /// XML tags to be enclosed by &lt;si&gt;&lt;/si&gt;...&lt;/si&gt; within the shared stings document</remarks>
        void AddFormattedValue(StringBuilder sb);

        
    }
}
