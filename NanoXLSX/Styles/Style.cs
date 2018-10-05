/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2018
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Runtime.CompilerServices;
using System.Text;
using NanoXLSX.Exceptions;

namespace Styles
{
    /// <summary>
    /// Class representing a Style with sub classes within a style sheet. An instance of this class is only a container for the different sub-classes. These sub-classes contain the actual styling information.
    /// </summary>
    public class Style : AbstractStyle
    {
        #region privateFields
        private string name;
        private bool internalStyle;
        private bool styleNameDefined;
        private StyleManager styleManagerReference;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the current Border object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public Border CurrentBorder { get; set; }
        /// <summary>
        /// Gets or sets the  current CellXf object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public CellXf CurrentCellXf { get; set; }
        /// <summary>
        /// Gets or sets the current Fill object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public Fill CurrentFill { get; set; }
        /// <summary>
        /// Gets or sets the  current Font object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public Font CurrentFont { get; set; }
        /// <summary>
        /// Gets or sets the  current NumberFormat object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public NumberFormat CurrentNumberFormat { get; set; }
        /// <summary>
        /// Gets or sets the name of the style. If not defined, the automatically calculated hash will be used as name
        /// </summary>
        [Append(Ignore = true)]
        public string Name
        {
            get { return name; }
            set
            {
                name = value;
                styleNameDefined = true;
            }
        }

        /// <summary>
        /// Sets the reference of the style manager
        /// </summary>
        [Append(Ignore = true)]
        public StyleManager StyleManagerReference
        {
            set
            {
                styleManagerReference = value;
                ReorganizeStyle();
            }
        }

        /// <summary>
        /// Gets whether the style is system internal. Such styles are not meant to be altered
        /// </summary>
        [Append(Ignore = true)]
        public bool IsInternalStyle
        {
            get { return internalStyle; }
        }

        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public Style()
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            styleNameDefined = false;
            name = CalculateHash();
        }

        /// <summary>
        /// Constructor with parameters
        /// </summary>
        /// <param name="name">Name of the style</param>
        public Style(string name)
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            styleNameDefined = false;
            this.name = name;
        }

        /// <summary>
        /// Constructor with parameters (internal use)
        /// </summary>
        /// <param name="name">Name of the style</param>
        /// <param name="forcedOrder">Number of the style for sorting purpose. Style will be placed to this position (internal use only)</param>
        /// <param name="internalStyle">If true, the style is marked as internal</param>
        public Style(string name, int forcedOrder, bool internalStyle)
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            this.name = name;
            InternalID = forcedOrder;
            this.internalStyle = internalStyle;
            styleNameDefined = true;
        }
        #endregion

        #region methods

        /// <summary>
        /// Appends the specified style parts to the current one. The parts can be instances of sub-classes like Border or CellXf or a Style instance. Only the altered properties of the specified style or style part that differs from a new / untouched style instance will be appended. This enables method chaining
        /// </summary>
        /// <param name="styleToAppend">The style to append or a sub-class of Style</param>
        /// <returns>Current style with appended style parts</returns>
        public Style Append(AbstractStyle styleToAppend)
        {
            if (styleToAppend.GetType() == typeof(Border))
            {
                CurrentBorder.CopyProperties((Border)styleToAppend, new Border());
            }
            else if (styleToAppend.GetType() == typeof(CellXf))
            {
                CurrentCellXf.CopyProperties((CellXf)styleToAppend, new CellXf());
            }
            else if (styleToAppend.GetType() == typeof(Fill))
            {
                CurrentFill.CopyProperties((Fill)styleToAppend, new Fill());
            }
            else if (styleToAppend.GetType() == typeof(Font))
            {
                CurrentFont.CopyProperties((Font)styleToAppend, new Font());
            }
            else if (styleToAppend.GetType() == typeof(NumberFormat))
            {
                CurrentNumberFormat.CopyProperties((NumberFormat)styleToAppend, new NumberFormat());
            }
            else if (styleToAppend.GetType() == typeof(Style))
            {
                CurrentBorder.CopyProperties(((Style)styleToAppend).CurrentBorder, new Border());
                CurrentCellXf.CopyProperties(((Style)styleToAppend).CurrentCellXf, new CellXf());
                CurrentFill.CopyProperties(((Style)styleToAppend).CurrentFill, new Fill());
                CurrentFont.CopyProperties(((Style)styleToAppend).CurrentFont, new Font());
                CurrentNumberFormat.CopyProperties(((Style)styleToAppend).CurrentNumberFormat, new NumberFormat());
            }
            return this;
        }

        /// <summary>
        /// Method to reorganize / synchronize the components of this style
        /// </summary>
        private void ReorganizeStyle()
        {
            if (styleManagerReference == null) { return; }

            Style newStyle = styleManagerReference.AddStyle(this);
            CurrentBorder = newStyle.CurrentBorder;
            CurrentCellXf = newStyle.CurrentCellXf;
            CurrentFill = newStyle.CurrentFill;
            CurrentFont = newStyle.CurrentFont;
            CurrentNumberFormat = newStyle.CurrentNumberFormat;

            if (styleNameDefined == false)
            {
                name = CalculateHash();
            }
        }

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class instance</returns>
        public override string ToString()
        {
            return InternalID + "->" + Hash;
        }

        /// <summary>
        /// Override method to calculate the hash of this component
        /// </summary>
        /// <returns>Calculated hash as string</returns>
        public sealed override string CalculateHash()
        {
            StringBuilder sb = new StringBuilder();
            if (CurrentBorder == null || CurrentCellXf == null || CurrentFill == null || CurrentFont == null || CurrentNumberFormat == null)
            {
                throw new StyleException("MissingReferenceException", "The hash of the style could not be created because one or more components are missing as references");
            }
            sb.Append(StyleManager.STYLEPREFIX);
            if (InternalID.HasValue)
            {
                sb.Append(InternalID.Value);
                sb.Append(':');
            }
            sb.Append(CurrentBorder.CalculateHash());
            sb.Append(CurrentCellXf.CalculateHash());
            sb.Append(CurrentFill.CalculateHash());
            sb.Append(CurrentFont.CalculateHash());
            sb.Append(CurrentNumberFormat.CalculateHash());
            return sb.ToString();
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            if (CurrentBorder == null || CurrentCellXf == null || CurrentFill == null || CurrentFont == null || CurrentNumberFormat == null)
            {
                throw new StyleException("MissingReferenceException", "The style could not be copied because one or more components are missing as references");
            }
            Style copy = new Style();
            copy.CurrentBorder = CurrentBorder.CopyBorder();
            copy.CurrentCellXf = CurrentCellXf.CopyCellXf();
            copy.CurrentFill = CurrentFill.CopyFill();
            copy.CurrentFont = CurrentFont.CopyFont();
            copy.CurrentNumberFormat = CurrentNumberFormat.CopyNumberFormat();
            return copy;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public Style CopyStyle()
        {
            return (Style)Copy();
        }

        #endregion

    }
    #region doc
    /// <summary>
    /// Sub-namespace for style definitions, style handling and (static) basic styles
    /// </summary>
    [CompilerGenerated]
    class NamespaceDoc // This class is only for documentation purpose (Sandcastle)
    { }
    #endregion

}