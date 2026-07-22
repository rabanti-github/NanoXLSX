/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

namespace NanoXLSX.Internal
{
    /// <summary>
    /// Describes a relationship entry or part that could not be discovered in tolerant reader mode.
    /// </summary>
    internal sealed class RelationshipDiscoveryIssue
    {
        /// <summary>
        /// Gets the ZIP entry path of the affected relationship part.
        /// </summary>
        public string RelationshipPartPath { get; private set; }

        /// <summary>
        /// Gets the affected relationship identifier, or null if the complete relationship part is affected.
        /// </summary>
        public string RelationshipId { get; private set; }

        /// <summary>
        /// Gets the reason why discovery could not retain the relationship data.
        /// </summary>
        public string Reason { get; private set; }

        /// <summary>
        /// Initializes a discovery issue.
        /// </summary>
        /// <param name="relationshipPartPath">ZIP entry path of the affected relationship part</param>
        /// <param name="relationshipId">Affected relationship identifier, or null if the complete relationship part is affected</param>
        /// <param name="reason">Reason why discovery could not retain the relationship data</param>
        public RelationshipDiscoveryIssue(string relationshipPartPath, string relationshipId, string reason)
        {
            RelationshipPartPath = relationshipPartPath;
            RelationshipId = relationshipId;
            Reason = reason;
        }
    }
}
