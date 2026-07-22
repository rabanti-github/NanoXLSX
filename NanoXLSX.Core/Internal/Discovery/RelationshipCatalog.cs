/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Holds the ordered relationship graph discovered while loading an OOXML package.
    /// </summary>
    internal sealed class RelationshipCatalog
    {
        private readonly List<RelationshipInfo> relationships = new List<RelationshipInfo>();
        private readonly List<RelationshipDiscoveryIssue> issues = new List<RelationshipDiscoveryIssue>();
        private readonly Dictionary<string, Dictionary<string, RelationshipInfo>> relationshipsBySource
            = new Dictionary<string, Dictionary<string, RelationshipInfo>>(StringComparer.Ordinal);
        private readonly ReadOnlyCollection<RelationshipInfo> readOnlyRelationships;
        private readonly ReadOnlyCollection<RelationshipDiscoveryIssue> readOnlyIssues;

        /// <summary>
        /// Gets all successfully discovered relationships in deterministic discovery order.
        /// </summary>
        public IReadOnlyList<RelationshipInfo> Relationships { get { return readOnlyRelationships; } }

        /// <summary>
        /// Gets all issues recorded while discovery was operating in tolerant mode.
        /// </summary>
        public IReadOnlyList<RelationshipDiscoveryIssue> Issues { get { return readOnlyIssues; } }

        /// <summary>
        /// Gets whether discovery completed without skipping an invalid relationship entry or part.
        /// </summary>
        public bool IsComplete { get { return issues.Count == 0; } }

        /// <summary>
        /// Initializes an empty relationship catalog.
        /// </summary>
        public RelationshipCatalog()
        {
            readOnlyRelationships = relationships.AsReadOnly();
            readOnlyIssues = issues.AsReadOnly();
        }

        /// <summary>
        /// Adds a relationship unless the same source part already contains its identifier.
        /// </summary>
        /// <param name="relationship">Relationship info object (cannot be null)</param>
        /// <returns>True if the relationship was added; otherwise false.</returns>
        public bool TryAdd(RelationshipInfo relationship)
        {
            string sourcePartPath = relationship.SourcePartPath ?? string.Empty;
            if (!relationshipsBySource.TryGetValue(sourcePartPath, out Dictionary<string, RelationshipInfo> sourceRelationships))
            {
                sourceRelationships = new Dictionary<string, RelationshipInfo>(StringComparer.Ordinal);
                relationshipsBySource.Add(sourcePartPath, sourceRelationships);
            }
            if (sourceRelationships.ContainsKey(relationship.Id))
            {
                return false;
            }
            sourceRelationships.Add(relationship.Id, relationship);
            relationships.Add(relationship);
            return true;
        }

        /// <summary>
        /// Adds an issue encountered during tolerant discovery.
        /// </summary>
        /// <param name="issue">Issue object (cannot be null)</param>
        public void AddIssue(RelationshipDiscoveryIssue issue)
        {
            issues.Add(issue);
        }

        /// <summary>
        /// Gets a relationship by its source part and source-local identifier.
        /// </summary>
        /// <param name="relationshipId">rID of the relationship</param>
        /// <param name="sourcePartPath">URI of the source path</param>
        /// <returns>Returns the relationship info object</returns>
        public RelationshipInfo GetBySourceAndId(string sourcePartPath, string relationshipId)
        {
            string normalizedSourcePartPath = sourcePartPath ?? string.Empty;
            if (relationshipsBySource.TryGetValue(normalizedSourcePartPath, out Dictionary<string, RelationshipInfo> sourceRelationships)
                && relationshipId != null
                && sourceRelationships.TryGetValue(relationshipId, out RelationshipInfo relationship))
            {
                return relationship;
            }
            return null;
        }

        /// <summary>
        /// Gets all relationships whose type exactly matches the supplied URI.
        /// </summary>
        /// <returns>Read-only list of relationships by type</returns>
        public IReadOnlyList<RelationshipInfo> GetByType(string documentType)
        {
            List<RelationshipInfo> matches = new List<RelationshipInfo>();
            foreach (RelationshipInfo relationship in relationships)
            {
                if (string.Equals(relationship.Type, documentType, StringComparison.Ordinal))
                {
                    matches.Add(relationship);
                }
            }
            return matches.AsReadOnly();
        }
    }
}
