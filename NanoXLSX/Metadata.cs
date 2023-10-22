/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2023
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using NanoXLSX.Shared.Utils;
using System;
using System.Globalization;
using System.Reflection;
using FormatException = NanoXLSX.Shared.Exceptions.FormatException;

namespace NanoXLSX
{
    /// <summary>
    /// Class representing the metadata of a workbook
    /// </summary>
    public class Metadata
    {

        #region privateFields
        private string applicationVersion;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the application which created the workbook. Default is NanoXLSX
        /// </summary>
        public string Application { get; set; }
        /// <summary>
        /// Gets or sets the version of the creation application. Default is the library version of NanoXLSX. Use the format xxxxx.yyyyy (e.g. 1.0 or 55.9875) in case of a custom value.
        /// </summary>
        /// <remarks>Allowed values are null, empty and fractions from 0.0  to 99999.99999 (max. number of digits before and after the period is 5)</remarks>
        public string ApplicationVersion
        {
            get { return applicationVersion; }
            set
            {
                applicationVersion = value;
                CheckVersion();
            }
        }
        /// <summary>
        /// Gets or sets the category of the document. There are no predefined values or restrictions about the content of this field
        /// </summary>
        public string Category { get; set; }
        /// <summary>
        /// Gets or sets the company owning the document. This value is for organizational purpose. Add more than one manager by using the semicolon (;) between the words
        /// </summary>
        public string Company { get; set; }
        /// <summary>
        /// Gets or sets the status of the document. There are no predefined values or restrictions about the content of this field
        /// </summary>
        public string ContentStatus { get; set; }
        /// <summary>
        /// Gets or sets the creator of the workbook. Add more than one creator by using the semicolon (;) between the authors
        /// </summary>
        public string Creator { get; set; }
        /// <summary>
        /// Gets or sets the description of the document or comment about it
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Gets or sets the hyperlink base of the document.
        /// </summary>
        public string HyperlinkBase { get; set; }
        /// <summary>
        /// Gets or sets the keywords of the workbook. Separate particular keywords with semicolons (;)
        /// </summary>
        public string Keywords { get; set; }
        /// <summary>
        /// Gets or sets the responsible manager of the document. This value is for organizational purpose.
        /// </summary>
        public string Manager { get; set; }
        /// <summary>
        /// Gets or sets the subject of the workbook
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// Gets or sets the title of the workbook
        /// </summary>
        public string Title { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Metadata()
        {
            Application = "NanoXLSX";
            Version vi = Assembly.GetExecutingAssembly().GetName().Version;
            ApplicationVersion = ParseVersion(vi.Major, vi.Minor, vi.Revision, vi.Build);
        }
        #endregion

        #region methods
        /// <summary>
        /// Checks the format of the passed version string. Allowed values are null, empty and fractions from 0.0  to 99999.99999 (max. number of digits before and after the period is 5)
        /// </summary>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if the version string is malformed</exception>
        private void CheckVersion()
        {
            if (string.IsNullOrEmpty(applicationVersion))
            { return; }
            string[] split = applicationVersion.Split('.');
            bool state = true;
            if (split.Length != 2)
            { state = false; }
            else
            {
                if (split[1].Length < 1 || split[1].Length > 5)
                { state = false; }
                if (split[0].Length < 1 || split[0].Length > 5)
                { state = false; }
            }
            if (!state)
            {
                throw new FormatException("The format of the version in the metadata is wrong (" + applicationVersion + "). Should be in the format and a range from '0.0' to '99999.99999'");
            }
        }
        #endregion

        #region staticMethods
        /// <summary>
        /// Method to parse a common version (major.minor.revision.build) into the compatible format (major.minor). The minimum value is 0.0 and the maximum value is 99999.99999<br></br>
        /// The minor, revision and build number are joined if possible. If the number is too long, 
        /// the additional characters will be removed from the right side down to five characters (e.g. 785563 will be 78556)
        /// </summary>
        /// <param name="major">Major number from 0 to 99999</param>
        /// <param name="minor">Minor number</param>
        /// <param name="build">Build number</param>
        /// <param name="revision">Revision number</param>
        /// <returns>Formatted version number (e.g. 1.0 or 55.987)</returns>
        /// <exception cref="NanoXLSX.Shared.Exceptions.FormatException">Throws a FormatException if the major number is too long or one of the numbers is negative</exception>
        public static string ParseVersion(int major, int minor, int build, int revision)
        {
            if (major < 0 || minor < 0 || build < 0 || revision < 0)
            {
                throw new FormatException("The format of the passed version is wrong. No negative number allowed.");
            }
            if (major > 99999)
            {
                throw new FormatException("The major number may not be bigger than 99999. The passed value is " + major);
            }
            string leftPart = ParserUtils.ToString(major);
            string rightPart = ParserUtils.ToString(minor) + ParserUtils.ToString(build) + ParserUtils.ToString(revision);
            rightPart = rightPart.TrimEnd('0');
            if (rightPart == "")
            { rightPart = "0"; }
            else
            {
                if (rightPart.Length > 5)
                {
                    rightPart = rightPart.Substring(0, 5);
                }
            }
            return leftPart + "." + rightPart;
        }
        #endregion
    }
}
