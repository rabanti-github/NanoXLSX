﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Utils.Xml;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface, used by XML writer classes 
    /// </summary>
    internal interface IPlugInWriter : IPlugIn
    {
        /// <summary>
        /// Gets or replaces the workbook instance, defined by the constructor
        /// </summary>
        Workbook Workbook { get; set; }

        /// <summary>
        /// Gets the main XML element, that is gerenated by <see cref="IPlugIn.Execute"/>
        /// </summary>
        XmlElement XmlElement { get; }

        /// <summary>
        /// Initialization method
        /// </summary>
        /// <param name="baseWriter">Base writer instance that holds any information for this writer</param>
        void Init(IBaseWriter baseWriter);

    }
}
