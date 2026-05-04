/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Text;
using Docs.IndexGenerator.Models;

namespace Docs.IndexGenerator.Generation
{
    /// <summary>
    /// Generates an llms.txt file following the convention at https://llmstxt.org/.
    /// Layout: H1 project title, blockquote summary, then linked sections.
    /// </summary>
    internal static class LlmsTxtGenerator
    {
        public static string Build(RootConfig root, MetaPackageConfig meta, PluginConfig plugins)
        {
            var sb = new StringBuilder();

            sb.Append("# ").AppendLine(root.ProjectName);
            sb.AppendLine();
            sb.Append("> ").AppendLine(root.LlmsSummary);
            sb.AppendLine();

            sb.AppendLine("## Packages");
            sb.AppendLine();
            sb.Append("Meta package **").Append(meta.PackageName).Append(" v").Append(meta.Version)
              .AppendLine("** bundles all of the below.");
            sb.AppendLine();

            string metaUrl = !string.IsNullOrEmpty(meta.NuGetUrl)
                ? meta.NuGetUrl
                : root.RepositoryUrl;
            sb.Append("- [").Append(meta.PackageName).Append("](").Append(metaUrl).Append("): ")
              .AppendLine(meta.Description ?? string.Empty);

            List<DocEntry> nonBundledEntries = new List<DocEntry>();

            foreach (DocEntry e in plugins.Entries)
            {
                if (!e.Bundled)
                {
                    nonBundledEntries.Add(e);
                    continue;
                }
                string url = !string.IsNullOrEmpty(e.NuGetUrl)
                    ? e.NuGetUrl
                    : (e.Repository ?? string.Empty);
                string description = e.Description ?? string.Empty;
                string bundledTag = e.Bundled ? " (bundled)" : string.Empty;
                sb.Append("- [").Append(e.Id).Append("](").Append(url).Append("): ")
                  .Append(description).AppendLine(bundledTag);
            }

            if (nonBundledEntries.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Other available packages (not bundled in meta package):");
                sb.AppendLine();
                foreach (DocEntry e in nonBundledEntries)
                {
                    string url = !string.IsNullOrEmpty(e.NuGetUrl)
                        ? e.NuGetUrl
                        : (e.Repository ?? string.Empty);
                    string description = e.Description ?? string.Empty;
                    sb.Append("- [").Append(e.Id).Append("](").Append(url).Append("): ")
                      .Append(description).AppendLine();
                }
            }

            sb.AppendLine();
            sb.AppendLine("## API Documentation");
            sb.AppendLine();
            sb.Append("- Combined documentation portal: ").AppendLine(root.ApiDocsUrl);
            foreach (var e in plugins.Entries)
            {
                if (!string.IsNullOrEmpty(e.ApiDocsUrl))
                {
                    sb.Append("- ").Append(e.Id).Append(": ").AppendLine(e.ApiDocsUrl);
                }
            }

            sb.AppendLine();
            sb.AppendLine("## Source");
            sb.AppendLine();
            sb.Append("- Primary repository: ").AppendLine(root.RepositoryUrl);
            foreach (var e in plugins.Entries)
            {
                if (!string.IsNullOrEmpty(e.Repository) &&
                    !string.Equals(e.Repository, root.RepositoryUrl, StringComparison.OrdinalIgnoreCase))
                {
                    sb.Append("- ").Append(e.Id).Append(" (external): ").AppendLine(e.Repository);
                }
            }

            sb.AppendLine();
            sb.AppendLine("## Demos");
            sb.AppendLine();
            sb.Append("- Repository with running demo use cases: ").AppendLine(root.DemoRepositoryUrl);
            sb.Append("- Direct folder URL with use cases for ").Append(root.ProjectName).Append(": ").AppendLine(root.DemoRepositoryUseCaseUrl);

            sb.AppendLine();
            sb.AppendLine("## Notes");
            sb.AppendLine();
            sb.AppendLine("- Docs.IndexGenerator/ is a build utility and not part of the public API.");
            sb.AppendLine("- This file is generated automatically on build from Docs.IndexGenerator/Config/*.json — do not edit by hand.");

            return sb.ToString();
        }
    }
}
