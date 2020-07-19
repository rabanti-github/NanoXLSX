using NanoXLSX.Exceptions;
using Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.LowLevel
{
    public class StyleReaderContainer
    {

        private List<CellXf> cellXfs = new List<CellXf>();
        private List<NumberFormat> numberFormats = new List<NumberFormat>();
        private List<Style> styles = new List<Style>();
        private List<Border> borders = new List<Border>();
        private List<Fill> fills = new List<Fill>();
        private List<Font> fonts = new List<Font>();


        public void AddStyleComponent(AbstractStyle component)
        {
            Type t = component.GetType();
            if (t == typeof(CellXf))
            {
                cellXfs.Add(component as CellXf);
            }
            else if (t == typeof(NumberFormat))
            {
                numberFormats.Add(component as NumberFormat);
            }
            else if (t == typeof(Style))
            {
                styles.Add(component as Style);
            }
            else if (t == typeof(Border))
            {
                borders.Add(component as Border);
            }
            else if (t == typeof(Fill))
            {
                fills.Add(component as Fill);
            }
            else if (t == typeof(Font))
            {
                fonts.Add(component as Font);
            }
            else
            {
                throw new StyleException("StyleException", "The style definition of the type '" + t.ToString() + "' is unknown or not implemented yet");
            }
        }

        public CellXf GetCellXF(int index, bool retunNullOnFail = false)
        {
            return GetComponnet(typeof(CellXf), index, retunNullOnFail) as CellXf;
        }
        public NumberFormat GetNumberFormat(int index, bool retunNullOnFail = false)
        {
            return GetComponnet(typeof(NumberFormat), index, retunNullOnFail) as NumberFormat;
        }

        public Style GetStyle(int index, bool retunNullOnFail = false)
        {
            return GetComponnet(typeof(Style), index, retunNullOnFail) as Style;
        }

        public Border GetBorder(int index, bool retunNullOnFail = false)
        {
            return GetComponnet(typeof(Border), index, retunNullOnFail) as Border;
        }

        public Fill GetFill(int index, bool retunNullOnFail = false)
        {
            return GetComponnet(typeof(Fill), index, retunNullOnFail) as Fill;
        }
        public Font GetFont(int index, bool retunNullOnFail = false)
        {
            return GetComponnet(typeof(Font), index, retunNullOnFail) as Font;
        }


        public int GetNextCellXFId()
        {
            return cellXfs.Count + 1;
        }

        public int GetNextNumberFormatId()
        {
            return numberFormats.Count + 1;
        }

        public int GetNextStyleId()
        {
            return styles.Count + 1;
        }

        public int GetNextBorderId()
        {
            return borders.Count + 1;
        }

        public int GetNextFillId()
        {
            return fills.Count + 1;
        }

        public int GetNextFontId()
        {
            return fonts.Count + 1;
        }
        private AbstractStyle GetComponnet(Type type, int index, bool retunNullOnFail)
        {
            try
            {
                if (type == typeof(CellXf))
                {
                    return cellXfs[index];
                }
                else if (type == typeof(NumberFormat))
                {
                    return numberFormats[index];
                }
                else if (type == typeof(Style))
                {
                    return styles[index];
                }
                else if (type == typeof(Border))
                {
                    return borders[index];
                }
                else if (type == typeof(Fill))
                {
                    return fills[index];
                }
                else if (type == typeof(Font))
                {
                    return fonts[index];
                }
                else
                {
                    throw new StyleException("StyleException", "The style definition of the type '" + type.ToString() + "' is unknown or not implemented yet");
                }
            }
            catch(Exception ex)
            {
                if (retunNullOnFail)
                {
                    return null;
                }
                else
                {
                    throw new StyleException("StyleException", "The style definition could not be retrieved. Please see inner exception:", ex);
                }
            }
        }


    }
}
