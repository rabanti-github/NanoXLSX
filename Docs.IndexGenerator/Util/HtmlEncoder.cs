/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace Docs.IndexGenerator.Util
{
    internal static class HtmlEncoder
    {
        private const string LineBreakPlaceholder = "##BR##";

        public static string Escape(string s)
        {
            if (s == null)
            {
                return string.Empty;
            }
            s = s.Replace("<br>", LineBreakPlaceholder)
                 .Replace("<br/>", LineBreakPlaceholder)
                 .Replace("<br />", LineBreakPlaceholder);
            s = System.Web.HttpUtility.HtmlEncode(s);
            return s.Replace(LineBreakPlaceholder, "<br>");
        }
    }
}
