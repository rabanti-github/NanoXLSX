/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */


using System;
using System.IO;
using System.Text.Json;

namespace Docs.IndexGenerator
{
    internal record DocEntry(string id, string title, string path, string? description);
    internal record RootConfig(string projectName, string version, string? description);
    internal record PluginConfig(DocEntry[] entries);

    public class Program
    {
        static int Main(string[] args)
        {

            string rootConfigPath = "../Docs.IndexGenerator/Config/root-config.json";
            string pluginConfigPath = "../Docs.IndexGenerator/Config/plugin-config.json";
            string outDir = Path.Combine("..", "docs"); // default relative to project dir

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--rootConfig" && i + 1 < args.Length)
                {
                    rootConfigPath = args[++i];
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
            if (!File.Exists(pluginConfigPath))
            {
                Console.Error.WriteLine($"Config file not found: {pluginConfigPath}");
                return 2;
            }

            string rootJson = File.ReadAllText(rootConfigPath);
            string pluginJson = File.ReadAllText(pluginConfigPath);
            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
            RootConfig rootConfig;
            PluginConfig pluginConfig;
            try
            {
                rootConfig = JsonSerializer.Deserialize<RootConfig>(rootJson, options) ?? throw new Exception("Root config deserialized to null");
                pluginConfig = JsonSerializer.Deserialize<PluginConfig>(pluginJson, options) ?? throw new Exception("Plugin config deserialized to null");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Failed to parse config: " + ex.Message);
                return 3;
            }

            Directory.CreateDirectory(outDir);

            // index.html
            string indexHtml = $@"<!doctype html>
<html lang=""en"">
<head>
  <meta charset=""utf-8"">
  <meta name=""viewport"" content=""width=device-width,initial-scale=1"">
  <title>{EscapeHtml(rootConfig.projectName)} — Documentation</title>
  <link rel=""stylesheet"" href=""style.css"">
</head>
<body>
  <main>
    <header>
    <h1>
        <img src=""NanoXLSX.png""
             alt=""NanoXLSX Icon""
             style=""height:48px; vertical-align:middle; margin-right:10px;"">
        {EscapeHtml(rootConfig.projectName)}
    </h1>
    <hr>
      <p class=""version"">Version {EscapeHtml(rootConfig.version)}</p>
      <p class=""desc"">{EscapeHtml(rootConfig.description ?? "")}</p>
    </header>

    <section>
      <h2>Available documentation</h2>
      <ul class=""list"">
{GenerateListItems(pluginConfig)}
      </ul>
    </section>
  </main>
</body>
</html>";

            File.WriteAllText(Path.Combine(outDir, "index.html"), indexHtml);

            // style.css (very small)
            string css = @"body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial; color: #222; margin: 32px; }
main { max-width: 900px; margin: auto; }
header h1 { margin: 0; }
.version { color: #666; margin-top: 4px; }
.list { line-height: 1.8; }
.list a { color: #0366d6; text-decoration: none; }
.list a:hover { text-decoration: underline; }";
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
            foreach (var e in cfg.entries)
            {
                string href = $"{Uri.EscapeUriString(e.path)}/index.html";
                sb.AppendLine($"        <li><a href=\"{href}\"><strong>{EscapeHtml(e.title)}</strong></a> — {EscapeHtml(e.description ?? "")}</li>");
            }
            return sb.ToString();
        }

        static string EscapeHtml(string s) => System.Web.HttpUtility.HtmlEncode(s);
    }
}
