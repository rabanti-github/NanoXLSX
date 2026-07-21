using System.IO.Packaging;
using NanoXLSX.Internal;
using Xunit;

namespace NanoXLSX.Core.Test.Misc
{
    public class RelationshipCatalogTest
    {
        private const string WORKSHEET_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

        [Fact(DisplayName = "Relationship IDs are scoped to their source part")]
        public void SourceScopedRelationshipIdTest()
        {
            RelationshipCatalog catalog = new RelationshipCatalog();
            RelationshipInfo workbookRelationship = CreateRelationship("rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet1.xml");
            RelationshipInfo worksheetRelationship = CreateRelationship("rId1", "http://example.org/drawing", "xl/worksheets/sheet1.xml", "xl/drawings/drawing1.xml");

            Assert.True(catalog.TryAdd(workbookRelationship));
            Assert.True(catalog.TryAdd(worksheetRelationship));
            Assert.Equal(2, catalog.Relationships.Count);
            Assert.Same(workbookRelationship, catalog.GetBySourceAndId("xl/workbook.xml", "rId1"));
            Assert.Same(worksheetRelationship, catalog.GetBySourceAndId("xl/worksheets/sheet1.xml", "rId1"));
        }

        [Fact(DisplayName = "Duplicate relationship IDs retain the first source-local entry")]
        public void DuplicateRelationshipIdTest()
        {
            RelationshipCatalog catalog = new RelationshipCatalog();
            RelationshipInfo first = CreateRelationship("rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet1.xml");
            RelationshipInfo duplicate = CreateRelationship("rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet2.xml");

            Assert.True(catalog.TryAdd(first));
            Assert.False(catalog.TryAdd(duplicate));
            Assert.Single(catalog.Relationships);
            Assert.Same(first, catalog.GetBySourceAndId("xl/workbook.xml", "rId1"));
        }

        [Fact(DisplayName = "Relationship type lookup is ordinal and case-sensitive")]
        public void ExactDocumentTypeLookupTest()
        {
            RelationshipCatalog catalog = new RelationshipCatalog();
            RelationshipInfo relationship = CreateRelationship("rId1", WORKSHEET_TYPE, "xl/workbook.xml", "xl/worksheets/sheet1.xml");
            catalog.TryAdd(relationship);

            Assert.Single(catalog.GetByType(WORKSHEET_TYPE));
            Assert.Empty(catalog.GetByType(WORKSHEET_TYPE.ToUpperInvariant()));
        }

        [Fact(DisplayName = "Discovery issues mark a relationship catalog as incomplete")]
        public void DiscoveryIssueTest()
        {
            RelationshipCatalog catalog = new RelationshipCatalog();
            Assert.True(catalog.IsComplete);

            RelationshipDiscoveryIssue issue = new RelationshipDiscoveryIssue("xl/_rels/workbook.xml.rels", "rId1", "Duplicate relationship identifier");
            catalog.AddIssue(issue);

            Assert.False(catalog.IsComplete);
            Assert.Single(catalog.Issues);
            Assert.Same(issue, catalog.Issues[0]);
        }

        private static RelationshipInfo CreateRelationship(string id, string type, string sourcePartPath, string targetPath)
        {
            return new RelationshipInfo(
                id,
                type,
                targetPath,
                TargetMode.Internal,
                "xl/_rels/workbook.xml.rels",
                sourcePartPath,
                targetPath);
        }
    }
}
