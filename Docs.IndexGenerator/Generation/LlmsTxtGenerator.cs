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

            // H1 + blockquote (required by spec)
            sb.Append("# ").AppendLine(root.ProjectName);
            sb.AppendLine();
            sb.Append("> ").AppendLine(root.LlmsSummary);

            // Optional prose paragraph between blockquote and first H2
            if (!string.IsNullOrEmpty(root.LlmsBodyParagraph))
            {
                sb.AppendLine();
                sb.AppendLine(root.LlmsBodyParagraph);
            }

            // Installation section
            bool hasPm = !string.IsNullOrEmpty(meta.PmInstallCommand);
            bool hasCli = !string.IsNullOrEmpty(meta.CliInstallCommand);
            if (hasPm || hasCli)
            {
                sb.AppendLine();
                sb.AppendLine("## Installation");
                sb.AppendLine();
                if (hasPm)
                    sb.Append("- Install via NuGet (Package Manager): `").Append(meta.PmInstallCommand).AppendLine("`");
                if (hasCli)
                    sb.Append("- Install via .NET CLI: `").Append(meta.CliInstallCommand).AppendLine("`");
            }

            // Packages section — spec-compliant [name](url): description links
            sb.AppendLine();
            sb.AppendLine("## Packages");
            sb.AppendLine();

            string metaUrl = !string.IsNullOrEmpty(meta.NuGetUrl) ? meta.NuGetUrl : root.RepositoryUrl;
            sb.Append("- [").Append(meta.PackageName).Append("](").Append(metaUrl).Append("): ")
              .AppendLine(meta.Description ?? string.Empty);

            foreach (DocEntry e in plugins.Entries)
            {
                string url = !string.IsNullOrEmpty(e.NuGetUrl) ? e.NuGetUrl : (e.Repository ?? string.Empty);
                string description = e.Description ?? string.Empty;
                string tag = e.Bundled ? " (bundled)" : string.Empty;
                sb.Append("- [").Append(e.Id).Append("](").Append(url).Append("): ")
                  .Append(description).AppendLine(tag);
            }

            // Usage / quick-start section
            if (!string.IsNullOrEmpty(root.LlmsUsageSnippet))
            {
                string lang = root.LlmsUsageSnippetLanguage ?? "csharp";
                sb.AppendLine();
                sb.AppendLine("## Usage");
                sb.AppendLine();
                sb.Append("```").AppendLine(lang);
                sb.AppendLine(root.LlmsUsageSnippet);
                sb.AppendLine("```");
            }

            // API Documentation — spec-compliant [name](url): description links
            sb.AppendLine();
            sb.AppendLine("## API Documentation");
            sb.AppendLine();
            sb.Append("- [Combined documentation portal](").Append(root.ApiDocsUrl).AppendLine("): All NanoXLSX packages");
            foreach (DocEntry e in plugins.Entries)
            {
                if (!string.IsNullOrEmpty(e.ApiDocsUrl))
                {
                    string note = !string.IsNullOrEmpty(e.Description) ? e.Description : e.Id + " API reference";
                    sb.Append("- [").Append(e.Id).Append("](").Append(e.ApiDocsUrl).Append("): ").AppendLine(note);
                }
            }

            // Source — spec-compliant [name](url): description links
            sb.AppendLine();
            sb.AppendLine("## Source");
            sb.AppendLine();
            sb.Append("- [").Append(root.ProjectName).Append("](").Append(root.RepositoryUrl).AppendLine("): Primary repository (meta-package, docs)");

            // Deduplicate external repos (multiple entries may share the same repo URL)
            var seenRepos = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DocEntry e in plugins.Entries)
            {
                if (string.IsNullOrEmpty(e.Repository) ||
                    string.Equals(e.Repository, root.RepositoryUrl, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!seenRepos.Add(e.Repository))
                    continue;
                string repoName = !string.IsNullOrEmpty(e.RepositoryDisplayName) ? e.RepositoryDisplayName : e.Id;
                sb.Append("- [").Append(repoName).Append("](").Append(e.Repository).AppendLine("): Source code (external)");
            }

            // Optional section (spec-defined: these entries can be skipped in short contexts)
            sb.AppendLine();
            sb.AppendLine("## Optional");
            sb.AppendLine();
            sb.Append("- [Demo repository](").Append(root.DemoRepositoryUrl).AppendLine("): Running demo use cases");
            sb.Append("- [").Append(root.ProjectName).Append(" demo use cases](").Append(root.DemoRepositoryUseCaseUrl).AppendLine("): Direct folder with NanoXLSX examples");
            if (!string.IsNullOrEmpty(root.WikiUrl))
                sb.Append("- [Getting started](").Append(root.WikiUrl).AppendLine("): Wiki / getting started guide");

            return sb.ToString();
        }
    }
}