/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;

namespace Docs.IndexGenerator.Generation
{
    internal static class StyleCssGenerator
    {
        public static void Write(string outputDirectory, string templateDirectory)
        {
            string source = Path.Combine(templateDirectory, "style.css");
            string destination = Path.Combine(outputDirectory, "style.css");
            if (!File.Exists(source))
            {
                throw new FileNotFoundException($"Style template not found: {source}");
            }
            File.Copy(source, destination, overwrite: true);
        }
    }
}
