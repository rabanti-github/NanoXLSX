/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

#nullable enable
namespace Docs.IndexGenerator.Models
{
    internal record DocEntry(
        string Id,
        string Title,
        string Path,
        string? Description,
        string? Repository,
        string? RepositoryDisplayName,
        bool Bundled,
        string? ApiDocsUrl,
        string? NuGetUrl);

    internal record RootConfig(
        string ProjectName,
        string BaseDescription,
        string RootDescription,
        string ApiDocsUrl,
        string RepositoryUrl,
        string DemoRepositoryUrl,
        string DemoRepositoryUseCaseUrl,
        string LlmsSummary);

    internal record MetaPackageConfig(
        string PackageName,
        string Version,
        string? Description,
        string? NuGetUrl);

    internal record PluginConfig(DocEntry[] Entries);
}
#nullable disable
