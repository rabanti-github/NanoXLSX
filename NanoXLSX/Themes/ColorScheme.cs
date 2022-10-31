using NanoXLSX.Shared.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.Themes
{
    public class ColorScheme : IColorScheme
    {

        private readonly int id;
        public string Name { get; set; }

        public IColor Dark1 { get; set; }
        public IColor Light1 { get; set; }
        public IColor Dark2 { get; set; }
        public IColor Light2 { get; set; }
        public IColor Accent1 { get; set; }
        public IColor Accent2 { get; set; }
        public IColor Accent3 { get; set; }
        public IColor Accent4 { get; set; }
        public IColor Accent5 { get; set; }
        public IColor Accent6 { get; set; }
        public IColor HyperLink { get; set; }
        public IColor FollowedHyperlink { get; set; }

        public ColorScheme(int id)
        {
            this.id = id;
        }

        public int GetSchemeId()
        {
            return id;
        }
    }
}
