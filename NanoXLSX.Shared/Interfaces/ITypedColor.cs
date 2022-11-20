/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Interfaces;

namespace NanoXLSX.Shared.Interfaces
{
    /// <summary>
    /// Interface to represent typed color with a specific value, based on a <see cref="IColor"/>
    /// </summary>
    public interface ITypedColor<T> : IColor
    {
        /// <summary>
        /// Sets or gets the color vale of the type <typeparamref name="T"/>
        /// </summary>
        T ColorValue { get; set; }
    }
}
