/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;

namespace Styles
{ 

    /// <summary>
    /// Attribute designated to control the copying of style properties
    /// </summary>
    /// <seealso cref="System.Attribute" />
    public class AppendAttribute : Attribute
    {
        /// <summary>
        /// Indicates whether the property annotated with the attribute is ignored during the copying of properties
        /// </summary>
        /// <value>
        ///   <c>true</c> if ignored, otherwise <c>false</c>.
        /// </value>
        public bool Ignore { get; set; }

        /// <summary>
        /// Indicates whether the property annotated with the attribute is a nested property. Nested properties are ignored but during the copying of properties but can be broken down to its sub-properties
        /// </summary>
        /// <value>
        ///   <c>true</c> if a nested property, otherwise <c>false</c>.
        /// </value>
        public bool NestedProperty { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        public AppendAttribute()
        {
            Ignore = false;
            NestedProperty = false;
        }
    }
}