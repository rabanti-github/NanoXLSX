/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


using System;
using System.IO;
using System.Text.Json;

namespace Docs.IndexGenerator
{
#nullable enable
    internal record DocEntry(string Id, string Title, string Path, string? Description, string? Repository, string? RepositoryDisplayName, bool Bundled);
    internal record RootConfig(string ProjectName, string BaseDescription, string RootDescription);
    internal record MetaPackageConfig(string PackageName, string Version, string? Description);
    internal record PluginConfig(DocEntry[] Entries);
#nullable disable

    public class Program
    {
        static int Main(string[] args)
        {

            string rootConfigPath = "../Docs.IndexGenerator/Config/root-config.json";
            string metaPackageConfigPath = "../Docs.IndexGenerator/Config/meta-package-config.json";
            string pluginConfigPath = "../Docs.IndexGenerator/Config/plugin-config.json";
            string outDir = Path.Combine("..", "Docs.IndexGenerator", "Output");

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--rootConfig" && i + 1 < args.Length)
                {
                    rootConfigPath = args[++i];
                }
                else if (args[i] == "--metaPackageConfig" && i + 1 < args.Length)
                {
                    metaPackageConfigPath = args[++i];
                }
                else if (args[i] == "--pluginConfig" && i + 1 < args.Length)
                {
                    pluginConfigPath = args[++i];
                }
                else if (args[i] == "--out" && i + 1 < args.Length)
                {
                    outDir = args[++i];
                }
            }

            if (!File.Exists(rootConfigPath))
            {
                Console.Error.WriteLine($"Config file not found: {rootConfigPath}");
                return 2;
            }
            if (!File.Exists(metaPackageConfigPath))
            {
                Console.Error.WriteLine($"Config file not found: {metaPackageConfigPath}");
                return 2;
            }
            if (!File.Exists(pluginConfigPath))
            {
                Console.Error.WriteLine($"Config file not found: {pluginConfigPath}");
                return 2;
            }

            string rootJson = File.ReadAllText(rootConfigPath);
            string metPackageJson = File.ReadAllText(metaPackageConfigPath);
            string pluginJson = File.ReadAllText(pluginConfigPath);
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            MetaPackageConfig metaPackageConfig;
            RootConfig rootConfig;
            PluginConfig pluginConfig;
            try
            {
                rootConfig = JsonSerializer.Deserialize<RootConfig>(rootJson, options) ?? throw new Exception("Root config deserialized to null");
                metaPackageConfig = JsonSerializer.Deserialize<MetaPackageConfig>(metPackageJson, options) ?? throw new Exception("Meta-package config deserialized to null");
                pluginConfig = JsonSerializer.Deserialize<PluginConfig>(pluginJson, options) ?? throw new Exception("Plugin config deserialized to null");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Failed to parse config: " + ex.Message);
                return 3;
            }

            Directory.CreateDirectory(outDir);

            // index.html
            string indexHtml = $@"
<!doctype html>
<html lang=""en"">
    <head>
      <meta charset=""utf-8"">
      <meta name=""viewport"" content=""width=device-width,initial-scale=1"">
      <title>{EscapeHtml(rootConfig.ProjectName)} — Documentation</title>
      <link rel=""stylesheet"" href=""style.css"">
    </head>
    <body>
      <main>
        <header>
        <h1>
            <img src=""NanoXLSX.png""
                 alt=""NanoXLSX""
                 style=""height:48px; vertical-align:middle; margin-right:10px;"">
            {EscapeHtml(rootConfig.ProjectName)}
        </h1>

        <p>{EscapeHtml(rootConfig.BaseDescription)}</p>
        <p>{EscapeHtml(rootConfig.RootDescription)}</p>

        <hr>

          <h2>Meta Package v{EscapeHtml(metaPackageConfig.Version)}</h2>
            <section>
            {GenerateMetaPackageItem(metaPackageConfig, rootConfig)}
            <p>There is no documentation for the meta package. Please see the section <b>Dependency Package Documentation</b> for the complete API documentation.</p>
            </section>

          <p class=""version"">Version {EscapeHtml(metaPackageConfig.Version)}</p>
        </header>

        <hr>

        <section>
          <h2>Dependency Package Documentation</h2>
          <table class=""list"">
            <tr>
            <td>Package</td><td>Description</td><td>Bundled</td><td>Repository</td>
            </tr>
    {GenerateListItems(pluginConfig)}
          </table>
        </section>
      </main>
    </body>
</html>";

            File.WriteAllText(Path.Combine(outDir, "index.html"), indexHtml);

            // style.css
            string css = @"body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial; color: #222; margin: 32px; }
hr {
    margin: 2rem 0;
    border: none;
    border-top: 1px solid #e5e7eb;
}
main {
    max-width: 900px;
    margin: 2rem auto;
    padding: 0 1rem;
    font-family: system-ui, -apple-system, BlinkMacSystemFont, ""Segoe UI"", Roboto, sans-serif;
}

header h1 {
    margin: 0;
}

.version {
    display: inline-block;
    background-color: #f2f4f7;
    color: #555;
    font-size: 0.85rem;
    font-weight: 600;
    padding: 3px 8px;
    border-radius: 6px;
    border: 1px solid #d0d7de;
    margin-top: 8px;
}

table.list {
    width: 100%;
    border-collapse: collapse;
    margin-top: 1rem;
    font-size: 0.95rem;
}

table.list th,
table.list td {
    padding: 0.6rem 0.75rem;
    text-align: left;
    border-bottom: 1px solid #e5e7eb;
    vertical-align: top;
}

table.list th {
    font-weight: 600;
    background-color: #f9fafb;
    border-bottom: 2px solid #d0d7de;
}

table.list tr:hover {
    background-color: #f6f8fa;
}

table.list a {
    color: #0366d6;
    text-decoration: none;
}

table.list a:hover {
    text-decoration: underline;
}
";
            File.WriteAllText(Path.Combine(outDir, "style.css"), css);
            var assetDest = Path.Combine(outDir, "NanoXLSX.png");

            string assetSrc = "../Docs.IndexGenerator/Assets/NanoXLSX.png";
            if (File.Exists("../Docs.IndexGenerator/Assets/NanoXLSX.png"))
            {
                File.Copy(assetSrc, assetDest, overwrite: true);
            }


            Console.WriteLine($"Generated index.html and style.css in: {Path.GetFullPath(outDir)}");
            return 0;
        }

        static string GenerateListItems(PluginConfig cfg)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var e in cfg.Entries)
            {
                string repoUrl = Uri.EscapeUriString(e.Repository);
                string docUrl = $"{Uri.EscapeUriString(e.Path)}/index.html";
                sb.AppendLine("  <tr>");
                sb.AppendLine($"    <td><a href=\"{docUrl}\"><strong>{EscapeHtml(e.Title)}</strong></a></td>");
                sb.AppendLine($"    <td>{EscapeHtml(e.Description ?? "")}</td>");
                sb.AppendLine($"    <td>{EscapeHtml(e.Bundled.ToString())}</td>");
                sb.AppendLine($"    <td><a href=\"{repoUrl}\" target=\"_blank\" rel=\"noopener\">{EscapeHtml(e.RepositoryDisplayName)}</a></td>");
                sb.AppendLine("  </tr>");

            }
            return sb.ToString();
        }

        static string GenerateMetaPackageItem(MetaPackageConfig metaPackageConfig, RootConfig rootConfig)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("<ul class=\"list\">");
            string description = CreatePrefix(rootConfig, metaPackageConfig.Description);
            sb.AppendLine($"        <li><strong>{EscapeHtml(metaPackageConfig.PackageName)}</strong> — {EscapeHtml(description ?? "")}</li>");
            sb.AppendLine("</ul>");
            return sb.ToString();
        }

        static string CreatePrefix(RootConfig config, string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }
            if (input.StartsWith(config.BaseDescription))
            {
                return input[config.BaseDescription.Length..];
            }
            return input;
        }

        static string EscapeHtml(string s)
        {
            // Use a placeholder that will never appear in normal text
            const string BR = "##BR##";
            s = s.Replace("<br>", BR).Replace("<br/>", BR).Replace("<br />", BR);
            s = System.Web.HttpUtility.HtmlEncode(s);

            // Restore <br>
            return s.Replace(BR, "<br>");
        }
    }
}
