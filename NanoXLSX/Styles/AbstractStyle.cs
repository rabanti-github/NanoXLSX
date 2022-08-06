/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NanoXLSX.Exceptions;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class represents an abstract style component
    /// </summary>
    public abstract class AbstractStyle : IComparable<AbstractStyle>
    {
        /// <summary>
        /// Gets or sets the internal ID for sorting purpose in the Excel style document (nullable)
        /// </summary>
        [Append(Ignore = true)]
        public int? InternalID { get; set; }


        /// <summary>
        /// Abstract method to copy a component (dereferencing)
        /// </summary>
        /// <returns>Returns a copied component</returns>
        public abstract AbstractStyle Copy();

        /// <summary>
        /// Internal method to copy altered properties from a source object. The decision whether a property is copied is dependent on a untouched reference object
        /// </summary>
        /// <typeparam name="T">Style or sub-class of Style that extends AbstractStyle</typeparam>
        /// <param name="source">Source object with properties to copy</param>
        /// <param name="reference">Reference object to decide whether the properties from the source objects are altered or not</param>
        internal void CopyProperties<T>(T source, T reference) where T : AbstractStyle
        {
            if (source == null || GetType() != source.GetType() && GetType() != reference.GetType())
            {
                throw new StyleException("The objects of the source, target and reference for style appending are not of the same type");
            }
            PropertyInfo[] infos = GetType().GetProperties();
            PropertyInfo sourceInfo;
            PropertyInfo referenceInfo;
            IEnumerable<AppendAttribute> attributes;
            foreach (PropertyInfo info in infos)
            {
                attributes = (IEnumerable<AppendAttribute>)info.GetCustomAttributes(typeof(AppendAttribute));
                if (attributes.Any() && !HandleProperties(attributes))
                {
                    continue;
                }
                sourceInfo = source.GetType().GetProperty(info.Name);
                referenceInfo = reference.GetType().GetProperty(info.Name);
                if (!sourceInfo.GetValue(source).Equals(referenceInfo.GetValue(reference)))
                {
                    info.SetValue(this, sourceInfo.GetValue(source));
                }
            }
        }

        /// <summary>
        /// Method to check whether a property is considered or skipped 
        /// </summary>
        /// <param name="attributes">Collection of attributes to check</param>
        /// <returns>Returns false as soon a property of the collection is marked as ignored or nested</returns>
        private static bool HandleProperties(IEnumerable<AppendAttribute> attributes)
        {
            foreach (AppendAttribute attribute in attributes)
            {
                if (attribute.Ignore || attribute.NestedProperty)
                {
                    return false; // skip property
                }
            }
            return true;
        }

        /// <summary>
        /// Method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
        public int CompareTo(AbstractStyle other)
        {
            if (!InternalID.HasValue)
            {
                return -1;
            }
            else if (other == null || !other.InternalID.HasValue)
            {
                return 1;
            }
            else
            {
                return InternalID.Value.CompareTo(other.InternalID.Value);
            }
        }

        /// <summary>
        /// Method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>True if both objects are equal, otherwise false</returns>
        public bool Equals(AbstractStyle other)
        {
            if (other == null)
            {
                return false;
            }
            return this.GetHashCode() == other.GetHashCode();
        }

        /// <summary>
        /// Append a JSON property for debug purpose (used in the ToString methods) to the passed string builder
        /// </summary>
        /// <param name="sb">String builder</param>
        /// <param name="name">Property name</param>
        /// <param name="value">Property value</param>
        /// <param name="terminate">If true, no comma and newline will be appended</param>
        internal static void AddPropertyAsJson(StringBuilder sb, String name, object value, bool terminate = false)
        {
            sb.Append("\"").Append(name).Append("\": ");
            if (value == null)
            {
                sb.Append("\"\"");
            }
            else
            {
                sb.Append("\"").Append(value.ToString().Replace("\"", "\\\"")).Append("\"");
            }
            if (!terminate)
            {
                sb.Append(",\n");
            }
        }

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
            /// Indicates whether the property annotated with the attribute is a nested property. 
            /// Nested properties are ignored during the copying of properties but can be broken down to its sub-properties
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

}
