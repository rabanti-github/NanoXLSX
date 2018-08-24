/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using NanoXLSX.Exception;
using FormatException = NanoXLSX.Exception.FormatException;

namespace NanoXLSX.Style
{
    /// <summary>
    /// Class represents an abstract style component
    /// </summary>
    public abstract class AbstractStyle : IComparable<AbstractStyle>, IEquatable<AbstractStyle>
    {
        /// <summary>
        /// Gets the unique hash of the object
        /// </summary>
        [Append(Ignore = true)]
        public string Hash
        {
            get { return CalculateHash(); }
        }

        /// <summary>
        /// Gets or sets the internal ID for sorting purpose in the Excel style document (nullable)
        /// </summary>
        [Append(Ignore = true)]
        public int? InternalID { get; set; }

        /// <summary>
        /// Abstract method definition to calculate the hash of the component
        /// </summary>
        /// <returns>Returns the hash of the component as string</returns>
        public abstract string CalculateHash();

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
            if (this.GetType() != source.GetType() && this.GetType() != reference.GetType())
            {
                throw new StyleException("CopyPropertyException", "The objects of the source, target and reference for style appending are not of the same type");
            }
            bool ignore;
            PropertyInfo[] infos = this.GetType().GetProperties();
            PropertyInfo sourceInfo, referenceInfo;
            IEnumerable<AppendAttribute> attributes;
            foreach (PropertyInfo info in infos)
            {
                attributes = (IEnumerable< AppendAttribute>)info.GetCustomAttributes(typeof(AppendAttribute));
                if (attributes.Count() > 0)
                {
                    ignore = false;
                    foreach (AppendAttribute attribute in attributes)
                    {
                        if (attribute.Ignore == true || attribute.NestedProperty == true)
                        {
                            ignore = true;
                            break;
                        }
                    }
                    if (ignore == true) { continue; } // skip property
                }

                sourceInfo = source.GetType().GetProperty(info.Name);
                referenceInfo = reference.GetType().GetProperty(info.Name);
                if (sourceInfo.GetValue(source).Equals(referenceInfo.GetValue(reference)) == false)
                {
                    info.SetValue(this, sourceInfo.GetValue(source));
                }
            }
        }

        /// <summary>
        /// Method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>-1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.</returns>
        public int CompareTo(AbstractStyle other)
        {
            if (InternalID.HasValue == false) { return -1; }
            else if (other.InternalID.HasValue == false) { return 1; }
            else { return InternalID.Value.CompareTo(other.InternalID.Value); }
        }

        /// <summary>
        /// Method to compare two objects for sorting purpose
        /// </summary>
        /// <param name="other">Other object to compare with this object</param>
        /// <returns>True if both objects are equal, otherwise false</returns>
        public bool Equals(AbstractStyle other)
        {
            return Hash.Equals(other.Hash);
        }

        /// <summary>
        /// Method to cast values of the components to string values for the hash calculation (protected/internal static method)
        /// </summary>
        /// <param name="o">Value to cast</param>
        /// <param name="sb">StringBuilder reference to put the casted object in</param>
        /// <param name="delimiter">Delimiter character to append after the casted value</param>
        protected static void CastValue(object o, ref StringBuilder sb, char? delimiter)
        {
            if (o == null)
            {
                sb.Append('#');
            }
            else if (o.GetType() == typeof(bool))
            {
                if ((bool)o == true) { sb.Append(1); }
                else { sb.Append(0); }
            }
            else if (o.GetType() == typeof(int))
            {
                sb.Append((int)o);
            }
            else if (o.GetType() == typeof(double))
            {
                sb.Append((double)o);
            }
            else if (o.GetType() == typeof(float))
            {
                sb.Append((float)o);
            }
            else if (o.GetType() == typeof(string))
            {
                if (o.ToString() == "#")
                {
                    sb.Append("_#_");
                }
                else
                {
                    sb.Append((string)o);
                }
            }
            else if (o.GetType() == typeof(long))
            {
                sb.Append((long)o);
            }
            else if (o.GetType() == typeof(char))
            {
                sb.Append((char)o);
            }
            else
            {
                sb.Append(o);
            }
            if (delimiter.HasValue == true)
            {
                sb.Append(delimiter.Value);
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
                this.Ignore = false;
                this.NestedProperty = false;
            }
        }
    }

    /*  ************************************************************************************  */




}
