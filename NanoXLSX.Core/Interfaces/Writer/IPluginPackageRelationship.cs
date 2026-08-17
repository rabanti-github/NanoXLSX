/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.IO.Packaging;

namespace NanoXLSX.Interfaces.Writer
{
    /// <summary>
    /// Interface used by package registry plug-ins to define a relationship (.rels) owned by a registered package part
    /// </summary>
    internal interface IPluginPackageRelationship
    {
        /// <summary>
        /// Gets the XML identifier of the relationship (rId). The identifier must be unique within the owning package part
        /// </summary>
        string RelationshipId { get; }

        /// <summary>
        /// Gets the absolute URI that identifies the role of the relationship
        /// </summary>
        string RelationshipType { get; }

        /// <summary>
        /// Gets the URI of the relationship target
        /// </summary>
        string Target { get; }

        /// <summary>
        /// Gets whether the target is internal or external to the package
        /// </summary>
        TargetMode TargetMode { get; }
    }
}
