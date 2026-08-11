using NanoXLSX.Internal;
using Xunit;

namespace NanoXLSX.Test.Core
{
    public class FeatureSetTest
    {
        [Fact(DisplayName = "A new aggregate feature set has no features")]
        public void Constructor_CreatesEmptyAggregateFeatureSet()
        {
            FeatureSet featureSet = new FeatureSet();

            AssertFeatures(featureSet, 0, 0, 0, 0);
        }

        [Fact(DisplayName = "The formula factory creates a valid formula feature set")]
        public void CreateFormula_CreatesFormulaFeatureSet()
        {
            FeatureSet featureSet = FeatureSet.CreateFormula();

            AssertFeatures(featureSet, 1, 0, 0, 0);
        }

        [Fact(DisplayName = "The defined-name factory creates a valid defined-name feature set")]
        public void CreateDefinedName_CreatesDefinedNameFeatureSet()
        {
            FeatureSet featureSet = FeatureSet.CreateDefinedName();

            AssertFeatures(featureSet, 0, 1, 0, 0);
        }

        [Fact(DisplayName = "Adding and removing feature sets updates all counts through the parent hierarchy")]
        public void AddAndRemove_PropagatesAllFeatures()
        {
            FeatureSet root = new FeatureSet();
            FeatureSet parent = new FeatureSet();
            FeatureSet aggregate = new FeatureSet();
            FeatureSet formula = FeatureSet.CreateFormula();
            FeatureSet definedName = FeatureSet.CreateDefinedName();

            parent.Add(root);
            formula.SetFormulaFeatures(true, true);
            definedName.SetDefinedNameFeatures(true, true);
            formula.Add(aggregate);
            definedName.Add(aggregate);

            AssertFeatures(aggregate, 2, 1, 1, 2);
            AssertFeatures(parent, 0, 0, 0, 0);
            AssertFeatures(root, 0, 0, 0, 0);

            aggregate.Add(parent);

            AssertFeatures(parent, 2, 1, 1, 2);
            AssertFeatures(root, 2, 1, 1, 2);

            formula.Remove(aggregate);

            AssertFeatures(formula, 1, 0, 1, 1);
            AssertFeatures(aggregate, 1, 1, 0, 1);
            AssertFeatures(parent, 1, 1, 0, 1);
            AssertFeatures(root, 1, 1, 0, 1);

            definedName.Remove(aggregate);

            AssertFeatures(definedName, 1, 1, 0, 1);
            AssertFeatures(aggregate, 0, 0, 0, 0);
            AssertFeatures(parent, 0, 0, 0, 0);
            AssertFeatures(root, 0, 0, 0, 0);
        }

        [Theory(DisplayName = "Setting formula features updates local and parent feature values")]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void SetFormulaFeatures_UpdatesFeatureValues(bool containsDefinedName, bool containsExternalLink)
        {
            FeatureSet root = new FeatureSet();
            FeatureSet parent = new FeatureSet();
            FeatureSet formula = FeatureSet.CreateFormula();
            parent.Add(root);
            formula.Add(parent);

            // Start from the opposite state so every InlineData case exercises a transition.
            formula.SetFormulaFeatures(!containsDefinedName, !containsExternalLink);
            // Test transitions from false to true and from true to false.
            formula.SetFormulaFeatures(containsDefinedName, containsExternalLink);

            int definedNameFormulaCount = containsDefinedName ? 1 : 0;
            int externalLinkCount = containsExternalLink ? 1 : 0;
            AssertFeatures(formula, 1, 0, definedNameFormulaCount, externalLinkCount);
            AssertFeatures(parent, 1, 0, definedNameFormulaCount, externalLinkCount);
            AssertFeatures(root, 1, 0, definedNameFormulaCount, externalLinkCount);

            formula.SetFormulaFeatures(containsDefinedName, containsExternalLink);

            AssertFeatures(formula, 1, 0, definedNameFormulaCount, externalLinkCount);
            AssertFeatures(parent, 1, 0, definedNameFormulaCount, externalLinkCount);
            AssertFeatures(root, 1, 0, definedNameFormulaCount, externalLinkCount);
        }

        [Theory(DisplayName = "Setting defined-name features updates local and parent feature values")]
        [InlineData(false, false)]
        [InlineData(false, true)]
        [InlineData(true, false)]
        [InlineData(true, true)]
        public void SetDefinedNameFeatures_UpdatesFeatureValues(bool isFormula, bool containsExternalLink)
        {
            FeatureSet root = new FeatureSet();
            FeatureSet parent = new FeatureSet();
            FeatureSet definedName = FeatureSet.CreateDefinedName();
            parent.Add(root);
            definedName.Add(parent);

            // Start from the opposite state so every InlineData case exercises a transition.
            definedName.SetDefinedNameFeatures(!isFormula, !containsExternalLink);
            // Test transitions from false to true and from true to false.
            definedName.SetDefinedNameFeatures(isFormula, containsExternalLink);

            int formulaCount = isFormula ? 1 : 0;
            int externalLinkCount = containsExternalLink ? 1 : 0;
            AssertFeatures(definedName, formulaCount, 1, 0, externalLinkCount);
            AssertFeatures(parent, formulaCount, 1, 0, externalLinkCount);
            AssertFeatures(root, formulaCount, 1, 0, externalLinkCount);

            definedName.SetDefinedNameFeatures(isFormula, containsExternalLink);

            AssertFeatures(definedName, formulaCount, 1, 0, externalLinkCount);
            AssertFeatures(parent, formulaCount, 1, 0, externalLinkCount);
            AssertFeatures(root, formulaCount, 1, 0, externalLinkCount);
        }

        [Fact(DisplayName = "Copying a feature set preserves its values without retaining its parent")]
        public void Copy_PreservesValuesWithoutParent()
        {
            FeatureSet parent = new FeatureSet();
            FeatureSet original = FeatureSet.CreateFormula();
            original.SetFormulaFeatures(true, true);
            original.Add(parent);

            FeatureSet copy = original.Copy();

            Assert.NotSame(original, copy);
            AssertFeatures(copy, 1, 0, 1, 1);

            copy.SetFormulaFeatures(false, false);

            AssertFeatures(copy, 1, 0, 0, 0);
            AssertFeatures(original, 1, 0, 1, 1);
            AssertFeatures(parent, 1, 0, 1, 1);
        }

        [Theory(DisplayName = "Changing a formula cell to a non-formula type removes its feature contribution")]
        [InlineData(Cell.CellType.String)]
        [InlineData(Cell.CellType.Number)]
        [InlineData(Cell.CellType.Bool)]
        [InlineData(Cell.CellType.Date)]
        [InlineData(Cell.CellType.Time)]
        [InlineData(Cell.CellType.Empty)]
        [InlineData(Cell.CellType.Error)]
        public void CellFormulaTypeChange_RemovesFeatures(Cell.CellType targetType)
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");
            Cell cell = worksheet.Cells["A1"];

            Assert.Equal(1, workbook.Features.FormulaCount);
            Assert.Equal(1, workbook.Features.ExternalLinkCount);

            cell.DataType = targetType;

            Assert.Equal(targetType, cell.DataType);
            Assert.Null(cell.Formula);
            Assert.Equal(0, worksheet.Features.FormulaCount);
            Assert.Equal(0, worksheet.Features.ExternalLinkCount);
            Assert.Equal(0, workbook.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);
        }

        [Fact(DisplayName = "A cached formula error preserves formula and external-link features")]
        public void CellFormulaCachedError_PreservesFeatures()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCellFormula("[1]ExternalSheet!A1", "A1");
            Cell cell = worksheet.Cells["A1"];

            cell.Formula.CachedValue = Enums.Errors.FormulaError.Reference;
            cell.Formula.CachedValueType = Cell.CellType.Error;

            Assert.Equal(Cell.CellType.Formula, cell.DataType);
            Assert.Equal("[1]ExternalSheet!A1", cell.Formula.Expression);
            Assert.Equal(1, worksheet.Features.FormulaCount);
            Assert.Equal(1, worksheet.Features.ExternalLinkCount);
            Assert.Equal(1, workbook.Features.FormulaCount);
            Assert.Equal(1, workbook.Features.ExternalLinkCount);
        }

        [Fact(DisplayName = "Replacing formula metadata keeps feature counters balanced")]
        public void CellFormulaReplacement_UpdatesFeatures()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            worksheet.AddCellFormula("A1", "A1");
            Cell cell = worksheet.Cells["A1"];

            cell.Formula = new FormulaData("[1]ExternalSheet!A1");

            Assert.Equal(1, worksheet.Features.FormulaCount);
            Assert.Equal(1, workbook.Features.FormulaCount);
            Assert.Equal(1, worksheet.Features.ExternalLinkCount);
            Assert.Equal(1, workbook.Features.ExternalLinkCount);

            cell.Value = null;

            Assert.Equal(Cell.CellType.Empty, cell.DataType);
            Assert.Null(cell.Formula);
            Assert.Equal(0, worksheet.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.FormulaCount);
            Assert.Equal(0, workbook.Features.ExternalLinkCount);
        }

        [Fact(DisplayName = "Changing a defined-name formula value removes its resolved reference feature")]
        public void CellDefinedNameFormulaValueChange_UpdatesFeatures()
        {
            Workbook workbook = new Workbook("Sheet1");
            Worksheet worksheet = workbook.CurrentWorksheet;
            DefinedName definedName = workbook.AddDefinedNameConstant("NamedValue", 1);
            worksheet.AddCellReference(definedName, "A1");
            Cell cell = worksheet.Cells["A1"];

            Assert.Equal(1, worksheet.Features.DefinedNameFormulaCount);
            Assert.Equal(1, workbook.Features.DefinedNameFormulaCount);

            cell.Value = "A1";

            Assert.Null(cell.Formula.DefinedNameReference);
            Assert.Equal("A1", cell.Formula.Expression);
            Assert.Equal(0, worksheet.Features.DefinedNameFormulaCount);
            Assert.Equal(0, workbook.Features.DefinedNameFormulaCount);
            Assert.Equal(1, workbook.Features.DefinedNameCount);
            Assert.Equal(1, workbook.Features.FormulaCount);
        }

        private static void AssertFeatures(
            FeatureSet featureSet,
            int formulaCount,
            int definedNameCount,
            int definedNameFormulaCount,
            int externalLinkCount)
        {
            Assert.Equal(formulaCount, featureSet.FormulaCount);
            Assert.Equal(definedNameCount, featureSet.DefinedNameCount);
            Assert.Equal(definedNameFormulaCount, featureSet.DefinedNameFormulaCount);
            Assert.Equal(formulaCount - definedNameFormulaCount, featureSet.WorksheetFormulaCount);
            Assert.Equal(externalLinkCount, featureSet.ExternalLinkCount);
            Assert.Equal(formulaCount > 0, featureSet.ContainsFormulas);
            Assert.Equal(definedNameCount > 0, featureSet.ContainsDefinedNames);
            Assert.Equal(definedNameFormulaCount > 0, featureSet.ContainsDefinedNameFormulas);
            Assert.Equal(formulaCount - definedNameFormulaCount > 0, featureSet.ContainsWorksheetFormulas);
            Assert.Equal(externalLinkCount > 0, featureSet.ContainsExternalLinks);
        }
    }
}
