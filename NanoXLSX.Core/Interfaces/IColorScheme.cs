/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


namespace NanoXLSX.Interfaces
{
    /// <summary>
    /// Interface to represent a color scheme that consists of 12 colors (<see cref="IColor"/>)
    /// </summary>
    public interface IColorScheme
    {
        /// <summary>
        /// Dark 1 (dk1) color of the color scheme
        /// </summary>
        IColor Dark1 { get; set; }
        /// <summary>
        /// Light 1 (lt1) color of the color scheme
        /// </summary>
        IColor Light1 { get; set; }
        /// <summary>
        /// Dark 2 (dk2) color of the color scheme
        /// </summary>
        IColor Dark2 { get; set; }
        /// <summary>
        /// Light 2 (lt2) color of the color scheme
        /// </summary>
        IColor Light2 { get; set; }
        /// <summary>
        /// Accent 1 (accent1) color of the color scheme
        /// </summary>
        IColor Accent1 { get; set; }
        /// <summary>
        /// Accent 2 (accent2) color of the color scheme
        /// </summary>
        IColor Accent2 { get; set; }
        /// <summary>
        /// Accent 3 (accent3) color of the color scheme
        /// </summary>
        IColor Accent3 { get; set; }
        /// <summary>
        /// Accent 4 (accent4) color of the color scheme
        /// </summary>
        IColor Accent4 { get; set; }
        /// <summary>
        /// Accent 5 (accent5) color of the color scheme
        /// </summary>
        IColor Accent5 { get; set; }
        /// <summary>
        /// Accent 6 (accent6) color of the color scheme
        /// </summary>
        IColor Accent6 { get; set; }
        /// <summary>
        /// Hyperlink (hlink) color of the color scheme
        /// </summary>
        IColor Hyperlink { get; set; }
        /// <summary>
        /// Followed Hyperlink (folHlink) color of the color scheme
        /// </summary>
        IColor FollowedHyperlink { get; set; }
    }
}
