/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Xml;
using NanoXLSX.Interfaces;
using NanoXLSX.Interfaces.Reader;
using NanoXLSX.Registry;
using NanoXLSX.Registry.Attributes;
using NanoXLSX.Utils.Xml;
using IOException = NanoXLSX.Exceptions.IOException;

namespace NanoXLSX.Internal.Readers
{
    /// <summary>
    /// Discovers relationship information in an OOXML package before document readers execute.
    /// </summary>
    [NanoXlsxPlugIn(PlugInUUID = PlugInUUID.DiscoveryReader)]
    public class DiscoveryReader : IDiscoveryReader
    {
        private const string RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
        private ZipArchive archive;

        /// <summary>
        /// Gets or sets the workbook that receives the temporary discovery catalog.
        /// </summary>
        public Workbook Workbook { get; set; }

        /// <summary>
        /// Gets or sets the reader options used for discovery validation.
        /// </summary>
        public IOptions Options { get; set; }

        /// <summary>
        /// Initializes a new discovery reader.
        /// </summary>
        public DiscoveryReader()
        {
        }

        /// <summary>
        /// Initializes discovery with a caller-owned ZIP archive.
        /// </summary>
        public void Init(ZipArchive archive, Workbook workbook, IOptions readerOptions)
        {
            this.archive = archive;
            Workbook = workbook;
            Options = readerOptions;
        }

        /// <summary>
        /// Discovers all valid relationship parts and prepares the temporary relationship catalog.
        /// </summary>
        /// <exception cref="InvalidOperationException">Thrown when the reader was not initialized with a ZIP archive or workbook.</exception>
        /// <exception cref="IOException">Thrown when invalid relationship data is encountered in strict validation mode.</exception>
        public void Execute()
        {
            if (archive == null)
            {
                throw new InvalidOperationException("The discovery reader was not initialized with a ZIP archive.");
            }
            if (Workbook == null)
            {
                throw new InvalidOperationException("The discovery reader was not initialized with a workbook.");
            }
            RelationshipCatalog catalog = DiscoverRelationships();
            Workbook.AuxiliaryData.SetData(PlugInUUID.DiscoveryReader, PlugInUUID.DiscoveryCatalogEntity, catalog);
        }

        private RelationshipCatalog DiscoverRelationships()
        {
            RelationshipCatalog catalog = new RelationshipCatalog();
            List<ZipArchiveEntry> relationshipEntries = archive.Entries
                .Where(entry => IsPotentialRelationshipPart(entry.FullName))
                .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
                .ToList();

            foreach (ZipArchiveEntry entry in relationshipEntries)
            {
                if (!TryGetSourcePartPath(entry.FullName, out string sourcePartPath, out string pathError))
                {
                    HandlePartIssue(catalog, entry.FullName, pathError, null);
                    continue;
                }
                DiscoverRelationshipPart(entry, sourcePartPath, catalog);
            }
            return catalog;
        }

        private void DiscoverRelationshipPart(ZipArchiveEntry entry, string sourcePartPath, RelationshipCatalog catalog)
        {
            List<RelationshipInfo> partRelationships = new List<RelationshipInfo>();
            List<RelationshipDiscoveryIssue> partIssues = new List<RelationshipDiscoveryIssue>();
            HashSet<string> relationshipIds = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                using (System.IO.Stream stream = entry.Open())
                using (XmlReader reader = XmlReader.Create(stream, XmlStreamUtils.CreateSettings()))
                {
                    bool rootFound = false;
                    while (reader.Read())
                    {
                        if (reader.NodeType != XmlNodeType.Element)
                        {
                            continue;
                        }
                        if (!rootFound)
                        {
                            if (reader.Depth != 0 || reader.LocalName != "Relationships" || reader.NamespaceURI != RELATIONSHIPS_NAMESPACE)
                            {
                                throw new XmlException("The relationship part has an invalid root element.");
                            }
                            rootFound = true;
                            continue;
                        }
                        if (reader.Depth != 1 || reader.LocalName != "Relationship" || reader.NamespaceURI != RELATIONSHIPS_NAMESPACE)
                        {
                            throw new XmlException("The relationship part contains an invalid element.");
                        }

                        RelationshipInfo relationship = ParseRelationship(reader, entry.FullName, sourcePartPath, partIssues);
                        if (relationship == null)
                        {
                            continue;
                        }
                        if (!relationshipIds.Add(relationship.Id))
                        {
                            string reason = "The relationship identifier is duplicated within its source part.";
                            if (IsStrictValidation)
                            {
                                throw CreateDiscoveryException(entry.FullName, relationship.Id, reason, null);
                            }
                            partIssues.Add(new RelationshipDiscoveryIssue(entry.FullName, relationship.Id, reason));
                            continue;
                        }
                        partRelationships.Add(relationship);
                    }
                    if (!rootFound)
                    {
                        throw new XmlException("The relationship part does not contain a Relationships root element.");
                    }
                }
            }
            catch (IOException)
            {
                throw;
            }
            catch (Exception ex)
            {
                HandlePartIssue(catalog, entry.FullName, "The relationship part could not be parsed.", ex);
                return;
            }

            foreach (RelationshipInfo relationship in partRelationships)
            {
                if (!catalog.TryAdd(relationship))
                {
                    string reason = "The relationship identifier is duplicated within its source part.";
                    if (IsStrictValidation)
                    {
                        throw CreateDiscoveryException(entry.FullName, relationship.Id, reason, null);
                    }
                    catalog.AddIssue(new RelationshipDiscoveryIssue(entry.FullName, relationship.Id, reason));
                }
            }
            foreach (RelationshipDiscoveryIssue issue in partIssues)
            {
                catalog.AddIssue(issue);
            }
        }

        private RelationshipInfo ParseRelationship(XmlReader reader, string relationshipPartPath, string sourcePartPath, List<RelationshipDiscoveryIssue> issues)
        {
            string id = reader.GetAttribute("Id");
            string type = reader.GetAttribute("Type");
            string target = reader.GetAttribute("Target");
            string targetModeValue = reader.GetAttribute("TargetMode");
            string reason = ValidateRequiredAttributes(id, type, target);
            if (reason != null)
            {
                return HandleRelationshipIssue(relationshipPartPath, id, reason, issues, null);
            }

            try
            {
                XmlConvert.VerifyNCName(id);
            }
            catch (Exception ex)
            {
                return HandleRelationshipIssue(relationshipPartPath, id, "The relationship identifier is not a valid XML NCName.", issues, ex);
            }
            if (!Uri.TryCreate(type, UriKind.Absolute, out Uri unusedTypeUri))
            {
                return HandleRelationshipIssue(relationshipPartPath, id, "The relationship type is not an absolute URI.", issues, null);
            }
            if (!Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out Uri targetUri))
            {
                return HandleRelationshipIssue(relationshipPartPath, id, "The relationship target is not a valid URI reference.", issues, null);
            }

            TargetMode targetMode;
            if (string.IsNullOrEmpty(targetModeValue) || targetModeValue == "Internal")
            {
                targetMode = TargetMode.Internal;
            }
            else if (targetModeValue == "External")
            {
                targetMode = TargetMode.External;
            }
            else
            {
                return HandleRelationshipIssue(relationshipPartPath, id, "The relationship TargetMode is invalid.", issues, null);
            }

            string resolvedTargetPath = null;
            if (targetMode == TargetMode.Internal)
            {
                if (targetUri.IsAbsoluteUri)
                {
                    return HandleRelationshipIssue(relationshipPartPath, id, "An internal relationship target cannot be an absolute URI.", issues, null);
                }
                try
                {
                    Uri sourceUri = new Uri("/" + sourcePartPath, UriKind.Relative);
                    Uri resolvedTargetUri = PackUriHelper.ResolvePartUri(sourceUri, targetUri);
                    resolvedTargetPath = resolvedTargetUri.OriginalString.TrimStart('/');
                }
                catch (Exception ex)
                {
                    return HandleRelationshipIssue(relationshipPartPath, id, "The internal relationship target could not be resolved as an OPC part URI.", issues, ex);
                }
            }

            return new RelationshipInfo(id, type, target, targetMode, relationshipPartPath, sourcePartPath, resolvedTargetPath);
        }

        private RelationshipInfo HandleRelationshipIssue(string relationshipPartPath, string relationshipId, string reason, List<RelationshipDiscoveryIssue> issues, Exception innerException)
        {
            if (IsStrictValidation)
            {
                throw CreateDiscoveryException(relationshipPartPath, relationshipId, reason, innerException);
            }
            issues.Add(new RelationshipDiscoveryIssue(relationshipPartPath, relationshipId, reason));
            return null;
        }

        private void HandlePartIssue(RelationshipCatalog catalog, string relationshipPartPath, string reason, Exception innerException)
        {
            if (IsStrictValidation)
            {
                throw CreateDiscoveryException(relationshipPartPath, null, reason, innerException);
            }
            catalog.AddIssue(new RelationshipDiscoveryIssue(relationshipPartPath, null, reason));
        }

        private static string ValidateRequiredAttributes(string id, string type, string target)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                return "The relationship Id attribute is missing or empty.";
            }
            if (string.IsNullOrWhiteSpace(type))
            {
                return "The relationship Type attribute is missing or empty.";
            }
            if (string.IsNullOrWhiteSpace(target))
            {
                return "The relationship Target attribute is missing or empty.";
            }
            return null;
        }

        private static bool IsPotentialRelationshipPart(string path)
        {
            return !string.IsNullOrEmpty(path)
                && path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
                && (path.StartsWith("_rels/", StringComparison.Ordinal) || path.IndexOf("/_rels/", StringComparison.Ordinal) >= 0);
        }

        private static bool TryGetSourcePartPath(string relationshipPartPath, out string sourcePartPath, out string error)
        {
            sourcePartPath = null;
            error = null;
            try
            {
                Uri relationshipPartUri = PackUriHelper.CreatePartUri(new Uri("/" + relationshipPartPath, UriKind.Relative));
                if (!PackUriHelper.IsRelationshipPartUri(relationshipPartUri)
                    || !string.Equals(relationshipPartUri.OriginalString.TrimStart('/'), relationshipPartPath, StringComparison.Ordinal))
                {
                    error = "The ZIP entry path is not a valid OPC relationship-part path.";
                    return false;
                }
                Uri sourcePartUri = PackUriHelper.GetSourcePartUriFromRelationshipPartUri(relationshipPartUri);
                sourcePartPath = sourcePartUri.OriginalString.TrimStart('/');
                return true;
            }
            catch (Exception ex)
            {
                error = "The ZIP entry path is not a valid OPC relationship-part path: " + ex.Message;
                return false;
            }
        }

        private static IOException CreateDiscoveryException(string relationshipPartPath, string relationshipId, string reason, Exception innerException)
        {
            string relationshipContext = relationshipId == null ? string.Empty : " (Id '" + relationshipId + "')";
            string message = "The relationship part '" + relationshipPartPath + "'" + relationshipContext + " is invalid. " + reason;
            return innerException == null ? new IOException(message) : new IOException(message, innerException);
        }

        private bool IsStrictValidation
        {
            get
            {
                ReaderOptions readerOptions = Options as ReaderOptions;
                return readerOptions != null && readerOptions.EnforceStrictValidation;
            }
        }
    }
}
