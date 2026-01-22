/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way  
 * Copyright Raphael Stoeckli © 2026
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Text;
using NanoXLSX.Exceptions;
using NanoXLSX.Utils;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX.Styles
{
    /// <summary>
    /// Class representing an XF entry. The XF entry is used to make reference to other style instances like Border or Fill and for the positioning of the cell content
    /// </summary>
    public class CellXf : AbstractStyle
    {
        #region constants
        /// <summary>
        /// Default horizontal align value as constant
        /// </summary>
        public static readonly HorizontalAlignValue DefaultHorizontalAlignment = HorizontalAlignValue.None;
        /// <summary>
        /// Default text break value as constant
        /// </summary>
        public static readonly TextBreakValue DefaultAlignment = TextBreakValue.None;
        /// <summary>
        /// Default text direction value as constant
        /// </summary>
        public static readonly TextDirectionValue DefaultTextDirection = TextDirectionValue.Horizontal;
        /// <summary>
        /// Default vertical align value as constant
        /// </summary>
        public static readonly VerticalAlignValue DefaultVerticalAlignment = VerticalAlignValue.None;
        #endregion

        #region privateFields
        private int textRotation;
        private TextDirectionValue textDirection;
        private int indent;
        #endregion

        #region enums
        /// <summary>
        /// Enum for the horizontal alignment of a cell, used by the <see cref="CellXf"/> class
        /// </summary>
        public enum HorizontalAlignValue
        {
            /// <summary>Content will be aligned left</summary>
            Left,
            /// <summary>Content will be aligned in the center</summary>
            Center,
            /// <summary>Content will be aligned right</summary>
            Right,
            /// <summary>Content will fill up the cell</summary>
            Fill,
            /// <summary>justify alignment</summary>
            Justify,
            /// <summary>General alignment</summary>
            General,
            /// <summary>Center continuous alignment</summary>
            CenterContinuous,
            /// <summary>Distributed alignment</summary>
            Distributed,
            /// <summary>No alignment. The alignment will not be used in a style</summary>
            None,
        }

        /// <summary>
        /// Enum for text break options, used by the <see cref="CellXf"/> class
        /// </summary>
        public enum TextBreakValue
        {
            /// <summary>Word wrap is active</summary>
            WrapText,
            /// <summary>Text will be resized to fit the cell</summary>
            ShrinkToFit,
            /// <summary>Text will overflow in cell</summary>
            None,
        }

        /// <summary>
        /// Enum for the general text alignment direction, used by the <see cref="CellXf"/> class
        /// </summary>
        public enum TextDirectionValue
        {
            /// <summary>Text direction is horizontal (default)</summary>
            Horizontal,
            /// <summary>Text direction is vertical</summary>
            Vertical,
        }

        /// <summary>
        /// Enum for the vertical alignment of a cell, used by the <see cref="CellXf"/> class
        /// </summary>
        public enum VerticalAlignValue
        {
            /// <summary>Content will be aligned on the bottom (default)</summary>
            Bottom,
            /// <summary>Content will be aligned on the top</summary>
            Top,
            /// <summary>Content will be aligned in the center</summary>
            Center,
            /// <summary>justify alignment</summary>
            Justify,
            /// <summary>Distributed alignment</summary>
            Distributed,
            /// <summary>No alignment. The alignment will not be used in a style</summary>
            None,
        }
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
                TextDirection = TextDirectionValue.Horizontal;
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
            HorizontalAlign = DefaultHorizontalAlignment;
            Alignment = DefaultAlignment;
            textDirection = DefaultTextDirection;
            VerticalAlign = DefaultVerticalAlignment;
            Locked = true; // Default in Excel
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
            if (textDirection == TextDirectionValue.Vertical)
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
        /// A hash code for this instance, suitable to be used in hashing algorithms and data structures like a hash table. 
        /// </returns>
        public override int GetHashCode()
        {
            unchecked
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
            CellXf copy = new CellXf
            {
                HorizontalAlign = HorizontalAlign,
                Alignment = Alignment,
                TextDirection = TextDirection,
                TextRotation = TextRotation,
                VerticalAlign = VerticalAlign,
                ForceApplyAlignment = ForceApplyAlignment,
                Locked = Locked,
                Hidden = Hidden,
                Indent = Indent
            };
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

        #region staticMethods

        /// <summary>
        /// Converts a HorizontalAlignValue enum to its string representation
        /// </summary>
        internal static string GetHorizontalAlignName(HorizontalAlignValue align)
        {
            string output = "";
            switch (align)
            {
                case HorizontalAlignValue.Left: output = "left"; break;
                case HorizontalAlignValue.Center: output = "center"; break;
                case HorizontalAlignValue.Right: output = "right"; break;
                case HorizontalAlignValue.Fill: output = "fill"; break;
                case HorizontalAlignValue.Justify: output = "justify"; break;
                case HorizontalAlignValue.General: output = "general"; break;
                case HorizontalAlignValue.CenterContinuous: output = "centerContinuous"; break;
                case HorizontalAlignValue.Distributed: output = "distributed"; break;
            }
            return output;
        }

        /// <summary>
        /// Converts a string to its corresponding HorizontalAlignValue enum
        /// </summary>
        internal static HorizontalAlignValue GetHorizontalAlignEnum(string name)
        {
            switch (name)
            {
                case "left": return HorizontalAlignValue.Left;
                case "center": return HorizontalAlignValue.Center;
                case "right": return HorizontalAlignValue.Right;
                case "fill": return HorizontalAlignValue.Fill;
                case "justify": return HorizontalAlignValue.Justify;
                case "general": return HorizontalAlignValue.General;
                case "centerContinuous": return HorizontalAlignValue.CenterContinuous;
                case "distributed": return HorizontalAlignValue.Distributed;
                default:
                    return HorizontalAlignValue.None;
            }
        }

        /// <summary>
        /// Converts a VerticalAlignValue enum to its string representation
        /// </summary>
        internal static string GetVerticalAlignName(VerticalAlignValue align)
        {
            string output = "";
            switch (align)
            {
                case VerticalAlignValue.Bottom: output = "bottom"; break;
                case VerticalAlignValue.Top: output = "top"; break;
                case VerticalAlignValue.Center: output = "center"; break;
                case VerticalAlignValue.Justify: output = "justify"; break;
                case VerticalAlignValue.Distributed: output = "distributed"; break;
            }
            return output;
        }

        /// <summary>
        /// Converts a string to its corresponding VerticalAlignValue enum
        /// </summary>
        internal static VerticalAlignValue GetVerticalAlignEnum(string name)
        {
            switch (name)
            {
                case "bottom": return VerticalAlignValue.Bottom;
                case "top": return VerticalAlignValue.Top;
                case "center": return VerticalAlignValue.Center;
                case "justify": return VerticalAlignValue.Justify;
                case "distributed": return VerticalAlignValue.Distributed;
                default:
                    return VerticalAlignValue.None;
            }
        }

        #endregion

    }

}
