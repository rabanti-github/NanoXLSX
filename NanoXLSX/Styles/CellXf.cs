/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2019
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Exceptions;

namespace Styles
{
    /// <summary>
    /// Class representing an XF entry. The XF entry is used to make reference to other style instances like Border or Fill and for the positioning of the cell content
    /// </summary>
    public class CellXf : AbstractStyle
    {
        #region enums
        /// <summary>
        /// Enum for the horizontal alignment of a cell 
        /// </summary>
        public enum HorizontalAlignValue
        {
            /// <summary>Content will be aligned left</summary>
            left,
            /// <summary>Content will be aligned in the center</summary>
            center,
            /// <summary>Content will be aligned right</summary>
            right,
            /// <summary>Content will fill up the cell</summary>
            fill,
            /// <summary>justify alignment</summary>
            justify,
            /// <summary>General alignment</summary>
            general,
            /// <summary>Center continuous alignment</summary>
            centerContinuous,
            /// <summary>Distributed alignment</summary>
            distributed,
            /// <summary>No alignment. The alignment will not be used in a style</summary>
            none,
        }

        /// <summary>
        /// Enum for text break options
        /// </summary>
        public enum TextBreakValue
        {
            /// <summary>Word wrap is active</summary>
            wrapText,
            /// <summary>Text will be resized to fit the cell</summary>
            shrinkToFit,
            /// <summary>Text will overflow in cell</summary>
            none,
        }

        /// <summary>
        /// Enum for the general text alignment direction
        /// </summary>
        public enum TextDirectionValue
        {
            /// <summary>Text direction is horizontal (default)</summary>
            horizontal,
            /// <summary>Text direction is vertical</summary>
            vertical,
        }

        /// <summary>
        /// Enum for the vertical alignment of a cell 
        /// </summary>
        public enum VerticalAlignValue
        {
            /// <summary>Content will be aligned on the bottom (default)</summary>
            bottom,
            /// <summary>Content will be aligned on the top</summary>
            top,
            /// <summary>Content will be aligned in the center</summary>
            center,
            /// <summary>justify alignment</summary>
            justify,
            /// <summary>Distributed alignment</summary>
            distributed,
            /// <summary>No alignment. The alignment will not be used in a style</summary>
            none,
        }
        #endregion

        #region privateFields
        private int textRotation;
        private TextDirectionValue textDirection;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets whether the applyAlignment property (used to merge cells) will be defined in the XF entry of the style. If true, applyAlignment will be defined
        /// </summary>
        public bool ForceApplyAlignment { get; set; }
        /// <summary>
        /// Gets or sets whether the hidden property (used for protection or hiding of cells) will be defined in the XF entry of the style. If true, hidden will be defined
        /// </summary>
        public bool Hidden { get; set; }
        /// <summary>
        /// Gets or sets the horizontal alignment of the style
        /// </summary>
        public HorizontalAlignValue HorizontalAlign { get; set; }
        /// <summary>
        /// Gets or sets whether the locked property (used for locking / protection of cells or worksheets) will be defined in the XF entry of the style. If true, locked will be defined
        /// </summary>
        public bool Locked { get; set; }
        /// <summary>
        /// Gets or sets the text break options of the style
        /// </summary>
        public TextBreakValue Alignment { get; set; }
        /// <summary>
        /// Gets or sets the direction of the text within the cell
        /// </summary>
        public TextDirectionValue TextDirection
        {
            get { return textDirection; }
            set
            {
                textDirection = value;
                CalculateInternalRotation();
            }
        }
        /// <summary>
        /// Gets or sets the text rotation in degrees (from +90 to -90)
        /// </summary>
        public int TextRotation
        {
            get { return textRotation; }
            set
            {
                textRotation = value;
                TextDirection = TextDirectionValue.horizontal;
                CalculateInternalRotation();
            }
        }
        /// <summary>
        /// Gets or sets the vertical alignment of the style
        /// </summary>
        public VerticalAlignValue VerticalAlign { get; set; }
        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public CellXf()
        {
            HorizontalAlign = HorizontalAlignValue.none;
            Alignment = TextBreakValue.none;
            textDirection = TextDirectionValue.horizontal;
            VerticalAlign = VerticalAlignValue.none;
            textRotation = 0;
        }
        #endregion

        #region methods
        /// <summary>
        /// Method to calculate the internal text rotation. The text direction and rotation are handled internally by the text rotation value
        /// </summary>
        /// <returns>Returns the valid rotation in degrees for internal uses (LowLevel)</returns>
        /// <exception cref="FormatException">Throws a FormatException if the rotation angle (-90 to 90) is out of range</exception>
        public int CalculateInternalRotation()
        {
            if (textRotation < -90 || textRotation > 90)
            {
                throw new FormatException("The rotation value (" + textRotation.ToString() + "°) is out of range. Range is form -90° to +90°");
            }
            if (textDirection == TextDirectionValue.vertical)
            {
                return 255;
            }
            else
            {
                if (textRotation >= 0)
                {
                    return textRotation;
                }
                else
                {
                    return (90 - textRotation);
                }
            }
        }

        /// <summary>
        /// Override toString method
        /// </summary>
        /// <returns>String of a class instance</returns>
        public override string ToString()
        {
            return "StyleXF:" + this.GetHashCode();
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            int p = 269;
            int r = 1;
            r *= p + (int)this.HorizontalAlign;
            r *= p + (int)this.VerticalAlign;
            r *= p + (int)this.Alignment;
            r *= p + (int)this.TextDirection;
            r *= p + this.TextRotation;
            r *= p + (this.ForceApplyAlignment ? 0 : 1);
            r *= p + (this.Locked ? 0 : 1);
            r *= p + (this.Hidden ? 0 : 1);
            return r;
        }

        /// <summary>
        /// Method to copy the current object to a new one without casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public override AbstractStyle Copy()
        {
            CellXf copy = new CellXf();
            copy.HorizontalAlign = HorizontalAlign;
            copy.Alignment = Alignment;
            copy.TextDirection = TextDirection;
            copy.TextRotation = TextRotation;
            copy.VerticalAlign = VerticalAlign;
            copy.ForceApplyAlignment = ForceApplyAlignment;
            copy.Locked = Locked;
            copy.Hidden = Hidden;
            return copy;
        }

        /// <summary>
        /// Method to copy the current object to a new one with casting
        /// </summary>
        /// <returns>Copy of the current object without the internal ID</returns>
        public CellXf CopyCellXf()
        {
            return (CellXf)Copy();
        }


        #endregion

    }

}