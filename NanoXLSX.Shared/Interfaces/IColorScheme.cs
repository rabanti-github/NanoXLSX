/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2022
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


namespace NanoXLSX.Shared.Interfaces
{
    /// <summary>
    /// Interface to represent a color scheme that consists of 12 colors (<see cref="IColor"/>)
    /// </summary>
    public interface IColorScheme
    {
        IColor Dark1 { get; set; }
        IColor Light1 { get; set; }
        IColor Dark2 { get; set; }
        IColor Light2 { get; set; }
        IColor Accent1 { get; set; }
        IColor Accent2 { get; set; }
        IColor Accent3 { get; set; }
        IColor Accent4 { get; set; }
        IColor Accent5 { get; set; }
        IColor Accent6 { get; set; }
        IColor HyperLink { get; set; }
        IColor FollowedHyperlink { get; set; }

        int GetSchemeId();
    }
}
