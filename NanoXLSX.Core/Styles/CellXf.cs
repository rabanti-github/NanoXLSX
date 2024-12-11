/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2024
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Shared.Interfaces.Styles;
using NanoXLSX.Shared.Exceptions;
using NanoXLSX.Shared.Utils;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing an XF entry. The XF entry is used to make reference to other style instances like Border or Fill and for the positioning of the cell content
    /// </summary>
    public class CellXf : AbstractStyle, ICellXf
    {
        #region constants
        /// <summary>
        /// Default horizontal align value as constant
        /// </summary>
        public static readonly HorizontalAlignValue DEFAULT_HORIZONTAL_ALIGNMENT = HorizontalAlignValue.none;
        /// <summary>
        /// Default text break value as constant
        /// </summary>
        public static readonly TextBreakValue DEFAULT_ALIGNMENT = TextBreakValue.none;
        /// <summary>
        /// Default text direction value as constant
        /// </summary>
        public static readonly TextDirectionValue DEFAULT_TEXT_DIRECTION = TextDirectionValue.horizontal;
        /// <summary>
        /// Default vertical align value as constant
        /// </summary>
        public static readonly VerticalAlignValue DEFAULT_VERTICAL_ALIGNMENT = VerticalAlignValue.none;
        #endregion

        #region privateFields
        private int textRotation;
        private TextDirectionValue textDirection;
        private int indent;
        #endregion

        #region properties
        /// <summary>
        /// Gets or sets whether the applyAlignment property (used to merge cells) will be defined in the XF entry of the style. If true, applyAlignment will be defined
        /// </summary>
        [Append]
        public bool ForceApplyAlignment { get; set; }
        /// <summary>
        /// Gets or sets whether the hidden property (used for protection or hiding of cells) will be defined in the XF entry of the style. If true, hidden will be defined
        /// </summary>
        [Append]
        public bool Hidden { get; set; }
        /// <summary>
        /// Gets or sets the horizontal alignment of the style
        /// </summary>
        [Append]
        public HorizontalAlignValue HorizontalAlign { get; set; }
        /// <summary>
        /// Gets or sets whether the locked property (used for locking / protection of cells or worksheets) will be defined in the XF entry of the style. If true, locked will be defined
        /// </summary>
        [Append]
        public bool Locked { get; set; }
        /// <summary>
        /// Gets or sets the text break options of the style
        /// </summary>
        [Append]
        public TextBreakValue Alignment { get; set; }
        /// <summary>
        /// Gets or sets the direction of the text within the cell
        /// </summary>
        [Append]
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
        [Append]
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
        [Append]
        public VerticalAlignValue VerticalAlign { get; set; }

        /// <summary>
        /// Gets or sets the indentation in case of left, right or distributed alignment. If 0, no alignment is applied
        /// </summary>
        [Append]
        public int Indent
        {
            get => indent;
            set
            {
                if (value >= 0)
                {
                    indent = value;
                }
                else
                {
                    throw new StyleException("The indent value '" + value + "' is not valid. It must be >= 0");
                }
            }
        }

        #endregion

        #region constructors
        /// <summary>
        /// Default constructor
        /// </summary>
        public CellXf()
        {
            HorizontalAlign = DEFAULT_HORIZONTAL_ALIGNMENT;
            Alignment = DEFAULT_ALIGNMENT;
            textDirection = DEFAULT_TEXT_DIRECTION;
            VerticalAlign = DEFAULT_VERTICAL_ALIGNMENT;
            textRotation = 0;
            Indent = 0;
        }
        #endregion

        #region methods
        /// <summary>
        /// Method to calculate the internal text rotation. The text direction and rotation are handled internally by the text rotation value
        /// </summary>
        /// <returns>Returns the valid rotation in degrees for internal use (LowLevel)</returns>
        /// <exception cref="FormatException">Throws a FormatException if the rotation angle (-90 to 90) is out of range</exception>
        internal int CalculateInternalRotation()
        {
            if (textRotation < -90 || textRotation > 90)
            {
                throw new FormatException("The rotation value (" + ParserUtils.ToString(textRotation) + "°) is out of range. Range is form -90° to +90°");
            }
            if (textDirection == TextDirectionValue.vertical)
            {
                textRotation = 255;
                return textRotation;
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
            StringBuilder sb = new StringBuilder();
            sb.Append("\"StyleXF\": {\n");
            AddPropertyAsJson(sb, "HorizontalAlign", HorizontalAlign);
            AddPropertyAsJson(sb, "Alignment", Alignment);
            AddPropertyAsJson(sb, "TextDirection", TextDirection);
            AddPropertyAsJson(sb, "TextRotation", TextRotation);
            AddPropertyAsJson(sb, "VerticalAlign", VerticalAlign);
            AddPropertyAsJson(sb, "ForceApplyAlignment", ForceApplyAlignment);
            AddPropertyAsJson(sb, "Locked", Locked);
            AddPropertyAsJson(sb, "Hidden", Hidden);
            AddPropertyAsJson(sb, "Indent", Indent);
            AddPropertyAsJson(sb, "HashCode", this.GetHashCode(), true);
            sb.Append("\n}");
            return sb.ToString();
        }

        /// <summary>
        /// Returns a hash code for this instance.
        /// </summary>
        /// <returns>
        /// A hash code for this instance, suitable for use in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            int hashCode = 626307906;
            hashCode = hashCode * -1521134295 + ForceApplyAlignment.GetHashCode();
            hashCode = hashCode * -1521134295 + Hidden.GetHashCode();
            hashCode = hashCode * -1521134295 + HorizontalAlign.GetHashCode();
            hashCode = hashCode * -1521134295 + Locked.GetHashCode();
            hashCode = hashCode * -1521134295 + Alignment.GetHashCode();
            hashCode = hashCode * -1521134295 + TextDirection.GetHashCode();
            hashCode = hashCode * -1521134295 + TextRotation.GetHashCode();
            hashCode = hashCode * -1521134295 + VerticalAlign.GetHashCode();
            hashCode = hashCode * -1521134295 + Indent.GetHashCode();
            return hashCode;
        }

        /// <summary>
        /// Returns whether two instances are the same
        /// </summary>
        /// <param name="obj">Object to compare</param>
        /// <returns>True if this instance and the other are the same</returns>
        public override bool Equals(object obj)
        {
            return obj is CellXf xf &&
                   ForceApplyAlignment == xf.ForceApplyAlignment &&
                   Hidden == xf.Hidden &&
                   HorizontalAlign == xf.HorizontalAlign &&
                   Locked == xf.Locked &&
                   Alignment == xf.Alignment &&
                   TextDirection == xf.TextDirection &&
                   TextRotation == xf.TextRotation &&
                   VerticalAlign == xf.VerticalAlign &&
                   Indent == xf.Indent;
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
            copy.Indent = Indent;
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
