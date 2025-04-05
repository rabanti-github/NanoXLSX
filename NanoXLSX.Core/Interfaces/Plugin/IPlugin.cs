﻿/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to define classes that can be handles by extension packages (plug-ins)
    /// </summary>
    internal interface IPlugIn
    {
        void Execute();
    }
}
