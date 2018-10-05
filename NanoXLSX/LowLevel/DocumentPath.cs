/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.LowLevel
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
            Filename = filename;
            Path = path;
        }

        /// <summary>
        /// Method to return the full path of the document
        /// </summary>
        /// <returns>Full path</returns>
        public string GetFullPath()
        {
            if (Path == null) { return Filename; }
            if (Path == "") { return Filename; }
            if (Path[Path.Length - 1] == System.IO.Path.AltDirectorySeparatorChar || Path[Path.Length - 1] == System.IO.Path.DirectorySeparatorChar)
            {
                return System.IO.Path.AltDirectorySeparatorChar + Path + Filename;
            }

            return System.IO.Path.AltDirectorySeparatorChar + Path + System.IO.Path.AltDirectorySeparatorChar + Filename;
        }

    }
    
}