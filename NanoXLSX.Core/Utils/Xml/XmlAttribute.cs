/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using System.Linq;

namespace NanoXLSX.Utils.Xml
{
    /// <summary>
    /// Struct representing an internally used XML attribute
    /// </summary>
    public struct XmlAttribute
    {
        /// <summary>
        /// Name of the attribute (without prefix)
        /// </summary>
        public string Name { get; private set; }
        /// <summary>
        /// Attribute value as string
        /// </summary>
        public string Value { get; private set; }
        /// <summary>
        /// True if a prefix for the attribute was defined 
        /// </summary>
        public bool HasPrefix { get; private set; }
        /// <summary>
        /// Prefix of the attribute. If not defined, the prefix will be an empty string
        /// </summary>
        public string Prefix { get; private set; }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="name">Attribute name</param>
        /// <param name="value">Attribute value</param>
        /// <param name="prefix">Optional attribute prefix</param>

        internal XmlAttribute(string name, string value, string prefix = "")
        {
            this.Name = name;
            this.Value = value;
            this.Prefix = prefix;
            HasPrefix = !string.IsNullOrEmpty(prefix);
        }

        /// <summary>
        /// Method to create an attribute instance
        /// </summary>
        /// <param name="name">Attribute name</param>
        /// <param name="value">Attribute value</param>
        /// <param name="prefix">Optional attribute prefix</param>
        /// <returns>Attribute instance</returns>
        public static XmlAttribute CreateAttribute(string name, string value, string prefix = "")
        {
            return new XmlAttribute(name, value, prefix);
        }

        /// <summary>
        /// Method to create an empty attribute instance
        /// </summary>
        /// <param name="name">Attribute name</param>
        /// <param name="prefix">Optional attribute prefix</param>
        /// <returns>Attribute instance</returns>
        public static XmlAttribute CreateEmptyAttribute(string name, string prefix = "")
        {
            return new XmlAttribute(name, "", prefix);
        }

        /// <summary>
        /// Method to find an attribute in a given list by attribute name. It is assumed that there are no duplicates (attribute name)
        /// </summary>
        /// <param name="name">Attribute name</param>
        /// <param name="attributes">List of attributes</param>
        /// <returns>Attribute that matche sthe name, or null if no attribute was found</returns>
        public static XmlAttribute? FindAttribute(string name, HashSet<XmlAttribute> attributes)
        {
            if (attributes == null || attributes.Count == 0)
            {
                return null;
            }
            if (!attributes.Any(a => a.Name == name))
            {
                return null;
            }
            return attributes.Where(a => a.Name == name).FirstOrDefault();
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is XmlAttribute attribute &&
                   Name == attribute.Name &&
                   Value == attribute.Value &&
                   Prefix == attribute.Prefix;
        }

        /// <summary>
        /// Gets the hash code of the attribute
        /// </summary>
        /// <returns>Hash code of the attribute</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = 27885120;
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Value);
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Prefix);
                return hashCode;
            }
        }
    }
}
