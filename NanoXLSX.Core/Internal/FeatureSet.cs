namespace NanoXLSX.Internal
{
    /// <summary>
    /// Class to count and aggregate features on the level of cells, worksheets and workbooks.
    /// </summary>
    /// \remark <remarks>Internal use only, without checks. Do not tamper with the class or instances</remarks>
    internal sealed class FeatureSet
    {
        /// <summary>
        /// Enum of the type of the feature set
        /// </summary>
        private enum FeatureSetType
        {
            /// <summary>Feature set is between the root and leaf element</summary>
            Aggregate,
            /// <summary>Feature set is on a formula (leaf element)</summary>
            Formula,
            /// <summary>Feature set is on defined name instance (leaf element) </summary>
            DefinedName
        }

        private FeatureSetType type;
        private FeatureSet parent;

        internal int FormulaCount { get; private set; }

        /// <summary>
        /// Number of defined-name definitions in the feature set.
        /// </summary>
        internal int DefinedNameCount { get; private set; }

        /// <summary>
        /// Number of formulas that directly reference a resolved defined name.
        /// </summary>
        /// \remark <remarks>Defined names embedded within larger formula expressions are currently not resolved and are not counted.</remarks>
        internal int DefinedNameFormulaCount { get; private set; }

        /// <summary>
        /// Number of formulas that are on worksheets and their cells, but not on defined names. This may only be useful on <see cref="FeatureSetType.Aggregate"/>
        /// </summary>
        internal int WorksheetFormulaCount
        {
            get
            {
                return FormulaCount - DefinedNameFormulaCount;
            }
        }

        /// <summary>
        /// Number of external links in the feature set
        /// </summary>
        internal int ExternalLinkCount { get; private set; }

        /// <summary>
        /// If true, the feature set contains formulas (defined names and cells)
        /// </summary>
        internal bool ContainsFormulas => FormulaCount > 0;
        /// <summary>
        /// If true, the feature set contains defined names
        /// </summary>
        internal bool ContainsDefinedNames => DefinedNameCount > 0;
        /// <summary>
        /// If true, the feature set contains formulas in defined names
        /// </summary>
        internal bool ContainsDefinedNameFormulas => DefinedNameFormulaCount > 0;
        /// <summary>
        /// If true, the feature set contains formulas on worksheets and their cells
        /// </summary>
        internal bool ContainsWorksheetFormulas => WorksheetFormulaCount > 0;
        /// <summary>
        /// If true, the feature set contains external links (defined names and cells)
        /// </summary>
        internal bool ContainsExternalLinks => ExternalLinkCount > 0;


        /// <summary>
        /// Creates a feature set representing exactly one formula.
        /// </summary>
        internal static FeatureSet CreateFormula()
        {
            return new FeatureSet(FeatureSetType.Formula);
        }

        /// <summary>
        /// Creates a feature set representing exactly one defined name
        /// </summary>
        /// <returns></returns>
        internal static FeatureSet CreateDefinedName()
        {
            return new FeatureSet(FeatureSetType.DefinedName);
        }
        /// <summary>
        /// Creates an aggregate feature set. Intended for Workbook and Worksheet.
        /// </summary>
        internal FeatureSet()
        {
            type = FeatureSetType.Aggregate;
        }

        /// <summary>
        /// Creates a feature set representing a non-aggregate feature
        /// </summary>
        /// <param name="type">Type of the feature set</param>
        private FeatureSet(FeatureSetType type)
        {
            this.type = type;
            //this.isFormulaFeatureSet = isFormulaFeatureSet;

            if (type == FeatureSetType.Formula)
            {
                FormulaCount = 1;
            }
            else if (type == FeatureSetType.DefinedName)
            {
                DefinedNameCount = 1;
            }
        }

        /// <summary>
        /// Adds this feature set to the specified parent. The current counts are propagated to the complete parent hierarchy.
        /// </summary>
        internal void Add(FeatureSet parent)
        {
            parent.ApplyDelta(
                FormulaCount,
                DefinedNameCount,
                DefinedNameFormulaCount,
                ExternalLinkCount);

            this.parent = parent;
        }

        /// <summary>
        /// Removes this feature set from the specified parent.
        /// The current counts are subtracted from the complete parent hierarchy.
        /// </summary>
        internal void Remove(FeatureSet parent)
        {
            parent.ApplyDelta(
                -FormulaCount,
                -DefinedNameCount,
                -DefinedNameFormulaCount,
                -ExternalLinkCount);

            this.parent = null;
        }

        /// <summary>
        /// Updates the features of a single formula and propagates all changes to the parent hierarchy.
        /// </summary>
        /// <param name="containsDefinedName">If true, the feature for defined names will be set</param>
        /// <param name="containsExternalLink">If true, the feature for external links will be set</param>
        internal void SetFormulaFeatures(bool containsDefinedName, bool containsExternalLink)
        {
            int newDefinedNameFormulaCount = containsDefinedName ? 1 : 0;
            int newExternalLinkCount = containsExternalLink ? 1 : 0;

            int definedNameFormulaDelta =
                newDefinedNameFormulaCount - DefinedNameFormulaCount;

            int externalLinkDelta =
                newExternalLinkCount - ExternalLinkCount;

            if (definedNameFormulaDelta == 0 && externalLinkDelta == 0)
            {
                return;
            }

            if (parent != null)
            {
                parent.ApplyDelta(
                    0,
                    0,
                    definedNameFormulaDelta,
                    externalLinkDelta);
            }

            DefinedNameFormulaCount = newDefinedNameFormulaCount;
            ExternalLinkCount = newExternalLinkCount;
        }

        /// <summary>
        /// Updates the features of a single defined name and propagates all changes to the parent hierarchy.
        /// </summary>
        /// <param name="isFormula">If true, the feature for formulas will be set</param>
        /// <param name="containsExternalLink">If true, the feature for external links will be set</param>
        internal void SetDefinedNameFeatures(bool isFormula, bool containsExternalLink)
        {
            int newFormulaCount = isFormula ? 1 : 0;
            int formulaDelta = newFormulaCount - FormulaCount;

            int newExternalLinkCount = containsExternalLink ? 1 : 0;
            int externalLinkDelta = newExternalLinkCount - ExternalLinkCount;

            if (externalLinkDelta == 0 && formulaDelta == 0)
            {
                return;
            }

            if (parent != null)
            {
                parent.ApplyDelta(formulaDelta, 0, 0, externalLinkDelta);
            }

            FormulaCount = newFormulaCount;
            ExternalLinkCount = newExternalLinkCount;
        }

        /// <summary>
        /// Applies a count difference to this aggregate FeatureSet and recursively propagates it to its parent.
        /// </summary>
        /// <param name="formulaDelta">Delta value for formula count</param>
        /// <param name="definedNameDelta">Delta value for defined name count</param>
        /// <param name="definedNameFormulaDelta">Delta value for formulas referencing defined names</param>
        /// <param name="externalLinkDelta">Delta value for external link count</param>
        private void ApplyDelta(
            int formulaDelta,
            int definedNameDelta,
            int definedNameFormulaDelta,
            int externalLinkDelta)
        {
            FormulaCount += formulaDelta;
            DefinedNameCount += definedNameDelta;
            DefinedNameFormulaCount += definedNameFormulaDelta;
            ExternalLinkCount += externalLinkDelta;

            if (parent != null)
            {
                parent.ApplyDelta(
                    formulaDelta,
                    definedNameDelta,
                    definedNameFormulaDelta,
                    externalLinkDelta);
            }
        }

        /// <summary>
        /// Copies the instance with counters bot not with parent. 
        /// The parent will be added by regular Add functions (e.g. <see cref="Worksheet.AddNextCellFormula(string)"/>)
        /// </summary>
        /// <returns>Returns the copied instance</returns>
        internal FeatureSet Copy()
        {
            FeatureSet copy = new FeatureSet(type)
            {
                FormulaCount = FormulaCount,
                DefinedNameCount = DefinedNameCount,
                DefinedNameFormulaCount = DefinedNameFormulaCount,
                ExternalLinkCount = ExternalLinkCount
            };

            return copy;
        }
    }
}
