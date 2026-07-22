/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO.Packaging;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Describes one relationship discovered in an OOXML package.
    /// </summary>
    internal sealed class RelationshipInfo
    {
        /// <summary>
        /// Gets the relationship identifier. The identifier is unique only within <see cref="SourcePartPath"/>.
        /// </summary>
        public string Id { get; private set; }

        /// <summary>
        /// Gets the relationship type URI exactly as declared in the relationship part.
        /// </summary>
        public string Type { get; private set; }

        /// <summary>
        /// Gets the unmodified relationship target.
        /// </summary>
        public string Target { get; private set; }

        /// <summary>
        /// Gets whether the relationship points to an internal package part or an external resource.
        /// </summary>
        public TargetMode TargetMode { get; private set; }

        /// <summary>
        /// Gets the ZIP entry path of the relationship part.
        /// </summary>
        public string RelationshipPartPath { get; private set; }

        /// <summary>
        /// Gets the normalized path of the source part. An empty string identifies the package root.
        /// </summary>
        public string SourcePartPath { get; private set; }

        /// <summary>
        /// Gets the normalized ZIP entry path of an internal target, or null for an external target.
        /// </summary>
        public string ResolvedTargetPath { get; private set; }

        /// <summary>
        /// Initializes a relationship description.
        /// </summary>
        /// <param name="id">Relationship identifier</param>
        /// <param name="type">Relationship type URI</param>
        /// <param name="target">Unmodified relationship target</param>
        /// <param name="targetMode">Internal or external resource target</param>
        /// <param name="relationshipPartPath">ZIP entry path of the relationship part</param>
        /// <param name="sourcePartPath">Normalized path of the source part</param>
        /// <param name="resolvedTargetPath">Normalized ZIP entry path of an internal target, or null for an external target</param>
        public RelationshipInfo(string id, string type, string target, TargetMode targetMode, string relationshipPartPath, string sourcePartPath, string resolvedTargetPath)
        {
            Id = id;
            Type = type;
            Target = target;
            TargetMode = targetMode;
            RelationshipPartPath = relationshipPartPath;
            SourcePartPath = sourcePartPath;
            ResolvedTargetPath = resolvedTargetPath;
        }
    }
}
