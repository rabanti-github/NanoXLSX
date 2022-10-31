using NanoXLSX.Shared.Interfaces;
using NanoXLSX.Shared.Utils;

namespace NanoXLSX.Themes
{
    public class SrgbColor : ITypedColor<string>
    {
        private string colorValue;

        public string ColorValue { get => colorValue; set 
            {
                Validators.ValidateColor(value, false);
                colorValue = value; 
            }
        }

        public string StringValue => colorValue;

        public string ToArgbColor()
        {
            // Is already validated
            return "FF" + colorValue;
        }

    }
}
