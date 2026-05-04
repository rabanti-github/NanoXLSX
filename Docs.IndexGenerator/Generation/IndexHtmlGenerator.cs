/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Text;
using Docs.IndexGenerator.Models;
using Docs.IndexGenerator.Util;

namespace Docs.IndexGenerator.Generation
{
    internal static class IndexHtmlGenerator
    {
        public static string Build(RootConfig root, MetaPackageConfig meta, PluginConfig plugins)
        {
            return $@"
<!doctype html>
<html lang=""en"">
    <head>
      <meta charset=""utf-8"">
      <meta name=""viewport"" content=""width=device-width,initial-scale=1"">
      <title>{HtmlEncoder.Escape(root.ProjectName)} — Documentation</title>
      <link rel=""stylesheet"" href=""style.css"">
    </head>
    <body>
      <main>
        <header>
        <h1>
            <img src=""NanoXLSX.png""
                 alt=""NanoXLSX""
                 style=""height:48px; vertical-align:middle; margin-right:10px;"">
            {HtmlEncoder.Escape(root.ProjectName)}
        </h1>

        <p>{HtmlEncoder.Escape(root.BaseDescription)}</p>
        <p>{HtmlEncoder.Escape(root.RootDescription)}</p>

        <hr>

          <h2>Meta Package v{HtmlEncoder.Escape(meta.Version)}</h2>
            <section>
            {RenderMetaPackageItem(meta, root)}
            <p>There is no documentation for the meta package. Please see the section <b>Dependency Package Documentation</b> for the complete API documentation.</p>
            </section>

          <p class=""version"">Version {HtmlEncoder.Escape(meta.Version)}</p>
        </header>

        <hr>

        <section>
          <h2>Dependency Package Documentation</h2>
          <table class=""list"">
            <tr>
            <td>Package</td><td>Description</td><td>Bundled</td><td>Repository</td>
            </tr>
    {RenderListItems(plugins)}
          </table>
        </section>
      </main>
    </body>
</html>";
        }

        private static string RenderListItems(PluginConfig cfg)
        {
            var sb = new StringBuilder();
            foreach (var e in cfg.Entries)
            {
                string repoUrl = Uri.EscapeUriString(e.Repository ?? string.Empty);
                string docUrl = $"{Uri.EscapeUriString(e.Path)}/index.html";
                sb.AppendLine("  <tr>");
                sb.AppendLine($"    <td><a href=\"{docUrl}\"><strong>{HtmlEncoder.Escape(e.Title)}</strong></a></td>");
                sb.AppendLine($"    <td>{HtmlEncoder.Escape(e.Description ?? "")}</td>");
                sb.AppendLine($"    <td>{HtmlEncoder.Escape(e.Bundled.ToString())}</td>");
                sb.AppendLine($"    <td><a href=\"{repoUrl}\" target=\"_blank\" rel=\"noopener\">{HtmlEncoder.Escape(e.RepositoryDisplayName ?? string.Empty)}</a></td>");
                sb.AppendLine("  </tr>");
            }
            return sb.ToString();
        }

        private static string RenderMetaPackageItem(MetaPackageConfig meta, RootConfig root)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<ul class=\"list\">");
            string description = StripBaseDescriptionPrefix(root, meta.Description);
            sb.AppendLine($"        <li><strong>{HtmlEncoder.Escape(meta.PackageName)}</strong> — {HtmlEncoder.Escape(description ?? "")}</li>");
            sb.AppendLine("</ul>");
            return sb.ToString();
        }

        private static string StripBaseDescriptionPrefix(RootConfig root, string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }
            if (input.StartsWith(root.BaseDescription, StringComparison.Ordinal))
            {
                return input[root.BaseDescription.Length..];
            }
            return input;
        }
    }
}
