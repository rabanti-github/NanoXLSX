
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.Themes
{
    public class Theme
    {
        public ColorScheme Colors { get; set; }
        public int ID { get; set; }

        public Theme(int id)
        {
            this.ID = id;
        }

        internal static Theme GetDefaultTheme()
        {
            Theme theme = new Theme(0);
            theme.Colors = new ColorScheme(0);
            return theme;
        }
    }
}
