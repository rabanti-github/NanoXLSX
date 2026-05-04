/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.IO;
using System.Text.Json;
using Docs.IndexGenerator.Models;

namespace Docs.IndexGenerator.Loading
{
    internal sealed class ConfigMissingException : Exception
    {
        public ConfigMissingException(string message) : base(message) { }
    }

    internal sealed class ConfigParseException : Exception
    {
        public ConfigParseException(string message, Exception inner) : base(message, inner) { }
    }

    internal static class ConfigLoader
    {
        public static (RootConfig Root, MetaPackageConfig Meta, PluginConfig Plugins) Load(
            string rootConfigPath,
            string metaPackageConfigPath,
            string pluginConfigPath)
        {
            EnsureExists(rootConfigPath);
            EnsureExists(metaPackageConfigPath);
            EnsureExists(pluginConfigPath);

            var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

            try
            {
                var root = JsonSerializer.Deserialize<RootConfig>(File.ReadAllText(rootConfigPath), options)
                    ?? throw new InvalidOperationException("Root config deserialized to null");
                var meta = JsonSerializer.Deserialize<MetaPackageConfig>(File.ReadAllText(metaPackageConfigPath), options)
                    ?? throw new InvalidOperationException("Meta-package config deserialized to null");
                var plugins = JsonSerializer.Deserialize<PluginConfig>(File.ReadAllText(pluginConfigPath), options)
                    ?? throw new InvalidOperationException("Plugin config deserialized to null");
                return (root, meta, plugins);
            }
            catch (Exception ex) when (ex is JsonException || ex is InvalidOperationException)
            {
                throw new ConfigParseException("Failed to parse config: " + ex.Message, ex);
            }
        }

        private static void EnsureExists(string path)
        {
            if (!File.Exists(path))
            {
                throw new ConfigMissingException($"Config file not found: {path}");
            }
        }
    }
}
