/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using Docs.IndexGenerator.Generation;
using Docs.IndexGenerator.Loading;

namespace Docs.IndexGenerator
{
    public class Program
    {
        private const int ExitOk = 0;
        private const int ExitConfigMissing = 2;
        private const int ExitConfigParseFailed = 3;

        static int Main(string[] args)
        {
            var opts = CliOptions.Parse(args);

            try
            {
                var (root, meta, plugins) = ConfigLoader.Load(
                    opts.RootConfigPath,
                    opts.MetaPackageConfigPath,
                    opts.PluginConfigPath);

                Directory.CreateDirectory(opts.OutDir);

                string indexHtml = IndexHtmlGenerator.Build(root, meta, plugins);
                File.WriteAllText(Path.Combine(opts.OutDir, "index.html"), indexHtml);

                StyleCssGenerator.Write(opts.OutDir, opts.TemplateDir);
                CopyAsset(opts.AssetSrc, Path.Combine(opts.OutDir, "NanoXLSX.png"));

                string llmsTxt = LlmsTxtGenerator.Build(root, meta, plugins);
                File.WriteAllText(opts.LlmsOutPath, llmsTxt);

                Console.WriteLine($"Generated index.html and style.css in: {Path.GetFullPath(opts.OutDir)}");
                Console.WriteLine($"Generated llms.txt at: {Path.GetFullPath(opts.LlmsOutPath)}");
                return ExitOk;
            }
            catch (ConfigMissingException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return ExitConfigMissing;
            }
            catch (ConfigParseException ex)
            {
                Console.Error.WriteLine(ex.Message);
                return ExitConfigParseFailed;
            }
        }

        private static void CopyAsset(string source, string destination)
        {
            if (File.Exists(source))
            {
                File.Copy(source, destination, overwrite: true);
            }
        }

        private sealed class CliOptions
        {
            public string RootConfigPath { get; set; } = "../Docs.IndexGenerator/Config/root-config.json";
            public string MetaPackageConfigPath { get; set; } = "../Docs.IndexGenerator/Config/meta-package-config.json";
            public string PluginConfigPath { get; set; } = "../Docs.IndexGenerator/Config/plugin-config.json";
            public string OutDir { get; set; } = Path.Combine("..", "Docs.IndexGenerator", "Output");
            public string TemplateDir { get; set; } = Path.Combine("..", "Docs.IndexGenerator", "Templates");
            public string AssetSrc { get; set; } = "../Docs.IndexGenerator/Assets/NanoXLSX.png";
            public string LlmsOutPath { get; set; } = Path.Combine("..", "llms.txt");

            public static CliOptions Parse(string[] args)
            {
                var o = new CliOptions();
                for (int i = 0; i < args.Length; i++)
                {
                    if (i + 1 >= args.Length)
                    {
                        break;
                    }
                    switch (args[i])
                    {
                        case "--rootConfig": o.RootConfigPath = args[++i]; break;
                        case "--metaPackageConfig": o.MetaPackageConfigPath = args[++i]; break;
                        case "--pluginConfig": o.PluginConfigPath = args[++i]; break;
                        case "--out": o.OutDir = args[++i]; break;
                        case "--templates": o.TemplateDir = args[++i]; break;
                        case "--asset": o.AssetSrc = args[++i]; break;
                        case "--llmsOut": o.LlmsOutPath = args[++i]; break;
                    }
                }
                return o;
            }
        }
    }
}
