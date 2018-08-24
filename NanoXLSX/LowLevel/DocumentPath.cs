/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.LowLevel
{
    partial class LowLevel
    {
        /// <summary>
        /// Class to manage XML document paths
        /// </summary>
        public class DocumentPath
        {
            /// <summary>
            /// File name of the document
            /// </summary>
            public string Filename { get; set; }
            /// <summary>
            /// Path of the document
            /// </summary>
            public string Path { get; set; }

            /// <summary>
            /// Default constructor
            /// </summary>
            public DocumentPath()
            {
            }

            /// <summary>
            /// Constructor with defined file name and path
            /// </summary>
            /// <param name="filename">File name of the document</param>
            /// <param name="path">Path of the document</param>
            public DocumentPath(string filename, string path)
            {
                this.Filename = filename;
                this.Path = path;
            }

            /// <summary>
            /// Method to return the full path of the document
            /// </summary>
            /// <returns>Full path</returns>
            public string GetFullPath()
            {
                if (this.Path == null) { return this.Filename; }
                if (this.Path == "") { return this.Filename; }
                if (this.Path[this.Path.Length - 1] == System.IO.Path.AltDirectorySeparatorChar || this.Path[this.Path.Length - 1] == System.IO.Path.DirectorySeparatorChar)
                {
                    return System.IO.Path.AltDirectorySeparatorChar.ToString() + this.Path + this.Filename;
                }
                else
                {
                    return System.IO.Path.AltDirectorySeparatorChar.ToString() + this.Path + System.IO.Path.AltDirectorySeparatorChar.ToString() + this.Filename;
                }
            }

        }
    }
}