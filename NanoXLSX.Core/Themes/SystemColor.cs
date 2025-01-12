/*
 * NanoXLSX is a small .NET library to generate and read XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2025
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */

using System.Collections.Generic;
using NanoXLSX.Exceptions;
using NanoXLSX.Interfaces;
using NanoXLSX.Utils;
using static NanoXLSX.Themes.SystemColor;

namespace NanoXLSX.Themes
{
    /// <summary>
    /// Class representing a predefined system color for certain purposes or target areas in the UI
    /// </summary>
    public class SystemColor : ITypedColor<Value>
    {
        public enum Value
        {
            ///<summary>3D Dark System Color: Specifies a Dark shadow color for three-dimensional display elements</summary>
            ThreeDimensionalDarkShadow,
            ///<summary>3D Light System Color: Specifies a Light color for three-dimensional display elements</summary>
            ThreeDimensionalLight,
            ///<summary>Active Border System Color: Specifies an Active Window Border Color</summary>
            ActiveBorder,
            ///<summary>Active Caption System Color: Specifies the active window title bar color. In particular the left side color in the color gradient of an active window"s title bar if the gradient effect is enabled</summary>
            ActiveCaption,
            ///<summary>Application Workspace System Color: Specifies the Background color of multiple document interface (MDI) applications</summary>
            AppWorkspace,
            ///<summary>Background System Color: Specifies the desktop background color</summary>
            Background,
            ///<summary>Button Face System Color: Specifies the face color for three-dimensional display elements and for dialog box backgrounds</summary>
            ButtonFace,
            ///<summary>Button Highlight System Color: Specifies the highlight color for three-dimensional display elements</summary>
            ButtonHighlight,
            ///<summary>Button Shadow System Color: Specifies the shadow color for three-dimensional display elements (for edges facing away from the light source)</summary>
            ButtonShadow,
            ///<summary>Button Text System Color: Specifies the color of text on push buttons</summary>
            ButtonText,
            ///<summary>Caption Text System Color: Specifies the color of text in the caption, size box, and scroll bar arrow box</summary>
            CaptionText,
            ///<summary>Gradient Active Caption System Color: Specifies the right side color in the color gradient of an active window"s title bar</summary>
            GradientActiveCaption,
            ///<summary>Gradient Inactive Caption System Color:  Specifies the right side color in the color gradient of an inactive window"s title bar</summary>
            GradientInactiveCaption,
            ///<summary>Gray Text System Color: Specifies a grayed (disabled) text. This color is set to 0 if the current display driver does not support a solid gray color</summary>
            GrayText,
            ///<summary>Highlight System Color: Specifies the color of Item(s) selected in a control</summary>
            Highlight,
            ///<summary>Highlight Text System Color: Specifies the text color of item(s) selected in a control</summary>
            HighlightText,
            ///<summary>Hot Light System Color: Specifies the color for a hyperlink or hot-tracked item</summary>
            HotLight,
            ///<summary>Inactive Border System Color: Specifies the color of the Inactive window border</summary>
            InactiveBorder,
            ///<summary>Inactive Caption System Color: Specifies the color of the Inactive window caption. Specifies the left side color in the color gradient of an inactive window"s title bar if the gradient effect is enabled</summary>
            InactiveCaption,
            ///<summary>Inactive Caption Text System Color: Specifies the color of text in an inactive caption</summary>
            InactiveCaptionText,
            ///<summary>Info Back System Color: Specifies the background color for tool tip controls</summary>
            InfoBackground,
            ///<summary>Info Text System Color: Specifies the text color for tool tip controls</summary>
            InfoText,
            ///<summary>Menu System Color: Specifies the menu background color</summary>
            Menu,
            ///<summary>Menu Bar System Color: Specifies the background color for the menu bar when menus appear as flat menus</summary>
            MenuBar,
            ///<summary>Menu Highlight System Color: Specifies the color used to highlight menu items when the menu appears as a flat menu</summary>
            MenuHighlight,
            ///<summary>Menu Text System Color: Specifies the color of Text in menus</summary>
            MenuText,
            ///<summary>Scroll Bar System Color: Specifies the scroll bar gray area color</summary>
            ScrollBar,
            ///<summary>Window System Color: Specifies window background color</summary>
            Window,
            ///<summary>Window Frame System Color: Specifies the window frame color</summary>
            WindowFrame,
            ///<summary>Window Text System Color: Specifies the color of text in windows</summary>
            WindowText,
        }


        private string lastColor = "000000";

        /// <summary>
        /// Gets or sets the enum value of the system color
        /// </summary>
        public Value ColorValue { get; set; } = Value.WindowText;

        /// <summary>
        /// Gets the internal OOXML string value of the enum, defined in <see cref="ColorValue"/>
        /// </summary>
        public string StringValue { get { return MapValueToString(this.ColorValue); } }

        /// <summary>
        /// Color value that was last computed by the generating application
        /// </summary>
        public string LastColor
        {
            get => lastColor;
            set
            {
                Validators.ValidateColor(value, false);
                lastColor = value;
            }
        }

        /// <summary>
        /// Default constructor
        /// </summary>
        public SystemColor()
        {
        }

        /// <summary>
        /// Constructor with value as parameter
        /// </summary>
        /// <param name="value">Color value of the system color</param>
        public SystemColor(Value value) : this()
        {
            this.ColorValue = value;
        }

        /// <summary>
        /// Constructor with all parameters
        /// </summary>
        /// <param name="value">Color value of the system color</param>
        /// <param name="lastColor">Last computed value</param>
        public SystemColor(Value value, string lastColor) : this(value)
        {
            this.LastColor = lastColor;
        }

        /// <summary>
        /// Maps the enum value of the system color to the OOXML value
        /// </summary>
        /// <param name="value">Enum value</param>
        /// <returns>String value that can be placed in an XML document</returns>
        private static string MapValueToString(Value value)
        {
            switch (value)
            {
                case Value.ThreeDimensionalDarkShadow: return "3dDkShadow";
                case Value.ThreeDimensionalLight: return "3dLight";
                case Value.ActiveBorder: return "activeBorder";
                case Value.ActiveCaption: return "activeCaption";
                case Value.AppWorkspace: return "appWorkspace";
                case Value.Background: return "background";
                case Value.ButtonFace: return "btnFace";
                case Value.ButtonHighlight: return "btnHighlight";
                case Value.ButtonShadow: return "btnShadow";
                case Value.ButtonText: return "btnText";
                case Value.CaptionText: return "captionText";
                case Value.GradientActiveCaption: return "gradientActiveCaption";
                case Value.GradientInactiveCaption: return "gradientInactiveCaption";
                case Value.GrayText: return "grayText";
                case Value.Highlight: return "highlight";
                case Value.HighlightText: return "highlightText";
                case Value.HotLight: return "hotLight";
                case Value.InactiveBorder: return "inactiveBorder";
                case Value.InactiveCaption: return "inactiveCaption";
                case Value.InactiveCaptionText: return "inactiveCaptionText";
                case Value.InfoBackground: return "infoBk";
                case Value.InfoText: return "infoText";
                case Value.Menu: return "menu";
                case Value.MenuBar: return "menuBar";
                case Value.MenuHighlight: return "menuHighlight";
                case Value.MenuText: return "menuText";
                case Value.ScrollBar: return "scrollBar";
                case Value.Window: return "window";
                case Value.WindowFrame: return "windowFrame";
                case Value.WindowText: return "windowText";
                default:
                    throw new StyleException(value + " is not a valid system color value");
            }
        }

        /// <summary>
        /// Maps a OOXML string value (from an XML document) to the corresponding enum value
        /// </summary>
        /// <param name="value">OOXML string value</param>
        /// <returns>Enum value</returns>
        internal static Value MapStringToValue(string value)
        {
            switch (value)
            {
                case "3dDkShadow": return Value.ThreeDimensionalDarkShadow;
                case "3dLight": return Value.ThreeDimensionalLight;
                case "activeBorder": return Value.ActiveBorder;
                case "activeCaption": return Value.ActiveCaption;
                case "appWorkspace": return Value.AppWorkspace;
                case "background": return Value.Background;
                case "btnFace": return Value.ButtonFace;
                case "btnHighlight": return Value.ButtonHighlight;
                case "btnShadow": return Value.ButtonShadow;
                case "btnText": return Value.ButtonText;
                case "captionText": return Value.CaptionText;
                case "gradientActiveCaption": return Value.GradientActiveCaption;
                case "gradientInactiveCaption": return Value.GradientInactiveCaption;
                case "grayText": return Value.GrayText;
                case "highlight": return Value.Highlight;
                case "highlightText": return Value.HighlightText;
                case "hotLight": return Value.HotLight;
                case "inactiveBorder": return Value.InactiveBorder;
                case "inactiveCaption": return Value.InactiveCaption;
                case "inactiveCaptionText": return Value.InactiveCaptionText;
                case "infoBk": return Value.InfoBackground;
                case "infoText": return Value.InfoText;
                case "menu": return Value.Menu;
                case "menuBar": return Value.MenuBar;
                case "menuHighlight": return Value.MenuHighlight;
                case "menuText": return Value.MenuText;
                case "scrollBar": return Value.ScrollBar;
                case "window": return Value.Window;
                case "windowFrame": return Value.WindowFrame;
                case "windowText": return Value.WindowText;
                default:
                    throw new StyleException(value + " is not a valid system color value");
            }
        }

        public override bool Equals(object obj)
        {
            return obj is SystemColor color &&
                   ColorValue == color.ColorValue &&
                   LastColor == color.LastColor;
        }

        public override int GetHashCode()
        {
            int hashCode = 1425985453;
            hashCode = hashCode * -1521134295 + ColorValue.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(LastColor);
            return hashCode;
        }
    }
}
