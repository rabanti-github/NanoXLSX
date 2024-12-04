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
        public IColor Hyperlink { get; set; }
        public IColor FollowedHyperlink { get; set; }

        public ColorScheme()
        {
        }


        public override bool Equals(object obj)
        {
            return obj is ColorScheme scheme &&
                   Name == scheme.Name &&
                   EqualityComparer<IColor>.Default.Equals(Dark1, scheme.Dark1) &&
                   EqualityComparer<IColor>.Default.Equals(Light1, scheme.Light1) &&
                   EqualityComparer<IColor>.Default.Equals(Dark2, scheme.Dark2) &&
                   EqualityComparer<IColor>.Default.Equals(Light2, scheme.Light2) &&
                   EqualityComparer<IColor>.Default.Equals(Accent1, scheme.Accent1) &&
                   EqualityComparer<IColor>.Default.Equals(Accent2, scheme.Accent2) &&
                   EqualityComparer<IColor>.Default.Equals(Accent3, scheme.Accent3) &&
                   EqualityComparer<IColor>.Default.Equals(Accent4, scheme.Accent4) &&
                   EqualityComparer<IColor>.Default.Equals(Accent5, scheme.Accent5) &&
                   EqualityComparer<IColor>.Default.Equals(Accent6, scheme.Accent6) &&
                   EqualityComparer<IColor>.Default.Equals(Hyperlink, scheme.Hyperlink) &&
                   EqualityComparer<IColor>.Default.Equals(FollowedHyperlink, scheme.FollowedHyperlink);
        }

        public override int GetHashCode()
        {
            int hashCode = -1016302979;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Dark1);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Light1);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Dark2);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Light2);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent1);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent2);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent3);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent4);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent5);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Accent6);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(Hyperlink);
            hashCode = hashCode * -1521134295 + EqualityComparer<IColor>.Default.GetHashCode(FollowedHyperlink);
            return hashCode;
        }

    }
}
