using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.Styles
{
    public class StyleRepository
    {
        private readonly object lockObject = new object();

        private static StyleRepository instance;

        public static StyleRepository Instance
        {
            get
            {
                instance = instance ?? new StyleRepository();
                return instance;
            }
        }

        private Dictionary<int, Style> styles;

        public Dictionary<int, Style> Styles { get => styles; }

        private StyleRepository()
        {
            styles = new Dictionary<int, Style>();
        }

        public Style AddStyle(Style style)
        {
            lock (lockObject)
            {
                if (style == null)
                {
                    return null;
                }
                int hashCode = style.GetHashCode();
                if (!styles.ContainsKey(hashCode))
                {
                    styles.Add(hashCode, style);
                }
                return styles[hashCode];
            }
        }

        public void FlushStyles()
        {
            styles.Clear();
        }


    }
}
