/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Interfaces;

namespace NanoXLSX.Themes
{
    /// <summary>
    /// Class representing a color scheme, used in a theme
    /// </summary>
    public class ColorScheme : IColorScheme
    {
        /// <summary>
        /// Name of the color scheme
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Theme color that defines the dark1 (t1) attribute of a theme 
        /// </summary>
        public IColor Dark1 { get; set; }
        /// <summary>
        /// Theme color that defines the light1 (bg1) attribute of a theme 
        /// </summary>
        public IColor Light1 { get; set; }
        /// <summary>
        /// Theme color that defines the dark2 (t2) attribute of a theme 
        /// </summary>
        public IColor Dark2 { get; set; }
        /// <summary>
        /// Theme color that defines the light2 (bg2) attribute of a theme 
        /// </summary>
        public IColor Light2 { get; set; }
        /// <summary>
        /// Theme color that defines the accent1 attribute of a theme 
        /// </summary>
        public IColor Accent1 { get; set; }
        /// <summary>
        /// Theme color that defines the accent2 attribute of a theme 
        /// </summary>
        public IColor Accent2 { get; set; }
        /// <summary>
        /// Theme color that defines the accent3 attribute of a theme 
        /// </summary>
        public IColor Accent3 { get; set; }
        /// <summary>
        /// Theme color that defines the accent4 attribute of a theme 
        /// </summary>
        public IColor Accent4 { get; set; }
        /// <summary>
        /// Theme color that defines the accent5 attribute of a theme 
        /// </summary>
        public IColor Accent5 { get; set; }
        /// <summary>
        /// Theme color that defines the accent6 attribute of a theme 
        /// </summary>
        public IColor Accent6 { get; set; }
        /// <summary>
        /// Theme color that defines the hyperlink attribute of a theme 
        /// </summary>
        public IColor Hyperlink { get; set; }
        /// <summary>
        /// Theme color that defines the followedHyperlink attribute of a theme 
        /// </summary>
        public IColor FollowedHyperlink { get; set; }

        /// <summary>
        /// Default constructor
        /// </summary>
        /// \remarks<remarks>The constructor does not initialize any of the color properties. 
        /// A workbook may become invalid on saving, if any of the values are remaining null or undefined. 
        /// This has to be maintained manually after initialization</remarks>
        public ColorScheme()
        {
            // NoOp
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is ColorScheme scheme &&
                   Name == scheme.Name &&
                   EqualityComparer<IColor>.Default.Equals(Dark1, scheme.Dark1) &&
                   EqualityComparer<IColor>.Default.Equals(Light1, scheme.Light1) &&
                   EqualityComparer<IColor>.Default.Equals(Dark2, scheme.Dark2) &&
                   EqualityComparer<IColor>.Default.Equals(Light2, scheme.Light2) &&
                   EqualityComparer<IColor>.Default.Equals(Accent1, scheme.Accent1) &&
                   EqualityComparer<IColor>.Default.Equals(Accent2, scheme.Accent2) &&
                   EqualityComparer<IColor>.Default.Equals(Accent3, scheme.Accent3) &&
                   EqualityComparer<IColor>.Default.Equals(Accent4, scheme.Accent4) &&
                   EqualityComparer<IColor>.Default.Equals(Accent5, scheme.Accent5) &&
                   EqualityComparer<IColor>.Default.Equals(Accent6, scheme.Accent6) &&
                   EqualityComparer<IColor>.Default.Equals(Hyperlink, scheme.Hyperlink) &&
                   EqualityComparer<IColor>.Default.Equals(FollowedHyperlink, scheme.FollowedHyperlink);
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hashCode = -1016302979;
                hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Dark1);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Light1);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Dark2);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Light2);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent1);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent2);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent3);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent4);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent5);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent6);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Hyperlink);
                hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(FollowedHyperlink);
                return hashCode;
            }
        }

    }
}
