/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Exceptions;
using NanoXLSX.Utils;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing a Style with sub classes within a style sheet. An instance of this class is only a container for the different sub-classes. These sub-classes contain the actual styling information.
    /// </summary>
    public class Style : AbstractStyle
    {
        #region privateFields
        private bool internalStyle;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets the current Border object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public Border CurrentBorder { get; set; }
        /// <summary>
        /// Gets or sets the current CellXf object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public CellXf CurrentCellXf { get; set; }
        /// <summary>
        /// Gets or sets the current Fill object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public Fill CurrentFill { get; set; }
        /// <summary>
        /// Gets or sets the current Font object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public Font CurrentFont { get; set; }
        /// <summary>
        /// Gets or sets the current NumberFormat object of the style
        /// </summary>
        [Append(NestedProperty = true)]
        public NumberFormat CurrentNumberFormat { get; set; }
        /// <summary>
        /// Gets or sets the name of the informal style. If not defined, the automatically calculated hash will be used as name
        /// </summary>
        /// \remark <remarks>The name is informal and not considered as an identifier, when collecting all styles for a workbook</remarks>
        [Append(Ignore = true)]
        public string Name { get; set; }

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
            Name = ParserUtils.ToString(this.GetHashCode());
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
            this.Name = name;
        }

        /// <summary>
        /// Constructor with parameters (internal use)
        /// </summary>
        /// <param name="name">Name of the style</param>
        /// <param name="forcedOrder">Number of the style for sorting purpose. The style will be placed at this position (internal use only)</param>
        /// <param name="internalStyle">If true, the style is marked as internal</param>
        public Style(string name, int forcedOrder, bool internalStyle)
        {
            CurrentBorder = new Border();
            CurrentCellXf = new CellXf();
            CurrentFill = new Fill();
            CurrentFont = new Font();
            CurrentNumberFormat = new NumberFormat();
            this.Name = name;
            InternalID = forcedOrder;
            this.internalStyle = internalStyle;
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
            if (styleToAppend == null)
            {
                return this;
            }
            if (styleToAppend.GetType() == typeof(Border))
            {
                CurrentBorder.CopyProperties<Border>((Border)styleToAppend, new Border());
            }
            else if (styleToAppend.GetType() == typeof(CellXf))
            {
                CurrentCellXf.CopyProperties<CellXf>((CellXf)styleToAppend, new CellXf());
            }
            else if (styleToAppend.GetType() == typeof(Fill))
            {
                CurrentFill.CopyProperties<Fill>((Fill)styleToAppend, new Fill());
            }
            else if (styleToAppend.GetType() == typeof(Font))
            {
                CurrentFont.CopyProperties<Font>((Font)styleToAppend, new Font());
            }
            else if (styleToAppend.GetType() == typeof(NumberFormat))
            {
                CurrentNumberFormat.CopyProperties<NumberFormat>((NumberFormat)styleToAppend, new NumberFormat());
            }
            else if (styleToAppend.GetType() == typeof(Style))
            {
                CurrentBorder.CopyProperties<Border>(((Style)styleToAppend).CurrentBorder, new Border());
                CurrentCellXf.CopyProperties<CellXf>(((Style)styleToAppend).CurrentCellXf, new CellXf());
                CurrentFill.CopyProperties<Fill>(((Style)styleToAppend).CurrentFill, new Fill());
                CurrentFont.CopyProperties<Font>(((Style)styleToAppend).CurrentFont, new Font());
                CurrentNumberFormat.CopyProperties<NumberFormat>(((Style)styleToAppend).CurrentNumberFormat, new NumberFormat());
            }
            return this;
        }

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class instance</returns>
        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{\n\"Style\": {\n");
            AddPropertyAsJson(sb, "Name", Name);
            AddPropertyAsJson(sb, "HashCode", this.GetHashCode());
            sb.Append(CurrentBorder.ToString()).Append(",\n");
            sb.Append(CurrentCellXf.ToString()).Append(",\n");
            sb.Append(CurrentFill.ToString()).Append(",\n");
            sb.Append(CurrentFont.ToString()).Append(",\n");
            sb.Append(CurrentNumberFormat.ToString()).Append("\n}\n}");
            return sb.ToString();
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash table. 
        /// </returns>
        /// <exception cref="StyleException">MissingReferenceException - The hash of the style could not be created because one or more components are missing as references</exception>
        public override int GetHashCode()
        {
            if (CurrentBorder == null || CurrentCellXf == null || CurrentFill == null || CurrentFont == null || CurrentNumberFormat == null)
            {
                throw new StyleException("The hash of the style could not be created because one or more components are missing as references");
            }
            unchecked
            {
                int p = 241;
                int r = 1;
                r *= p + this.CurrentBorder.GetHashCode();
                r *= p + this.CurrentCellXf.GetHashCode();
                r *= p + this.CurrentFill.GetHashCode();
                r *= p + this.CurrentFont.GetHashCode();
                r *= p + this.CurrentNumberFormat.GetHashCode();
                return r;
            }
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            if (CurrentBorder == null || CurrentCellXf == null || CurrentFill == null || CurrentFont == null || CurrentNumberFormat == null)
            {
                throw new StyleException("The style could not be copied because one or more components are missing as references");
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
}
