/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;

namespace NanoXLSX.Interfaces.Reader
{
    /// <summary>
    /// Interface, used by XML queue reader classes 
    /// </summary>
    internal interface IPluginQueueReader : IPluginBaseReader
    {
    }
}
