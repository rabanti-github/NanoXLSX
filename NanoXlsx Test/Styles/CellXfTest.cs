using NanoXLSX.Exceptions;
using NanoXLSX.Styles;
using System;
using Xunit;

namespace NanoXLSX_Test.Styles
{
    public class CellXfTest
    {
        private CellXf exampleStyle;

        public CellXfTest()
        {
            exampleStyle = new CellXf();
            exampleStyle.Hidden = true;
            exampleStyle.Locked = true;
            exampleStyle.ForceApplyAlignment = true;
            exampleStyle.HorizontalAlign = CellXf.HorizontalAlignValue.left;
            exampleStyle.VerticalAlign = CellXf.VerticalAlignValue.center;
            exampleStyle.TextDirection = CellXf.TextDirectionValue.horizontal;
            exampleStyle.Alignment = CellXf.TextBreakValue.shrinkToFit;
            exampleStyle.TextRotation = 75;
            exampleStyle.Indent = 3;
        }

        [Theory(DisplayName = "Test of the get and set function of the Hidden property")]
        [InlineData(true)]
        [InlineData(false)]
        public void HiddenTest(bool value)
        {
            CellXf cellXf = new CellXf();
            Assert.False(cellXf.Hidden);
            cellXf.Hidden = value;
            Assert.Equal(value, cellXf.Hidden);
        }

        [Theory(DisplayName = "Test of the get and set function of the Locked property")]
        [InlineData(true)]
        [InlineData(false)]
        public void LockedTest(bool value)
        {
            CellXf cellXf = new CellXf();
            Assert.True(cellXf.Locked); // Locked is set to true by default (has no effect until protection is enabled)
            cellXf.Locked = value;
            Assert.Equal(value, cellXf.Locked);
        }

        [Theory(DisplayName = "Test of the get and set function of the ForceApplyAlignment property")]
        [InlineData(true)]
        [InlineData(false)]
        public void ForceApplyAlignmentTest(bool value)
        {
            CellXf cellXf = new CellXf();
            Assert.False(cellXf.ForceApplyAlignment);
            cellXf.ForceApplyAlignment = value;
            Assert.Equal(value, cellXf.ForceApplyAlignment);
        }

        [Theory(DisplayName = "Test of the get and set function of the HorizontalAlign property")]
        [InlineData(CellXf.HorizontalAlignValue.center)]
        [InlineData(CellXf.HorizontalAlignValue.centerContinuous)]
        [InlineData(CellXf.HorizontalAlignValue.distributed)]
        [InlineData(CellXf.HorizontalAlignValue.fill)]
        [InlineData(CellXf.HorizontalAlignValue.general)]
        [InlineData(CellXf.HorizontalAlignValue.justify)]
        [InlineData(CellXf.HorizontalAlignValue.left)]
        [InlineData(CellXf.HorizontalAlignValue.none)]
        [InlineData(CellXf.HorizontalAlignValue.right)]
        public void HorizontalAlignTest(CellXf.HorizontalAlignValue value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(CellXf.DEFAULT_HORIZONTAL_ALIGNMENT, cellXf.HorizontalAlign); // none is default
            cellXf.HorizontalAlign = value;
            Assert.Equal(value, cellXf.HorizontalAlign);
        }

        [Theory(DisplayName = "Test of the get and set function of the VerticalAlign property")]
        [InlineData(CellXf.VerticalAlignValue.bottom)]
        [InlineData(CellXf.VerticalAlignValue.center)]
        [InlineData(CellXf.VerticalAlignValue.distributed)]
        [InlineData(CellXf.VerticalAlignValue.justify)]
        [InlineData(CellXf.VerticalAlignValue.none)]
        [InlineData(CellXf.VerticalAlignValue.top)]
        public void VerticalAlignTest(CellXf.VerticalAlignValue value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(CellXf.DEFAULT_VERTICAL_ALIGNMENT, cellXf.VerticalAlign); // none is default
            cellXf.VerticalAlign = value;
            Assert.Equal(value, cellXf.VerticalAlign);
        }


        [Theory(DisplayName = "Test of the get and set function of the HorizontalAlign property")]
        [InlineData(CellXf.TextDirectionValue.horizontal)]
        [InlineData(CellXf.TextDirectionValue.vertical)]
        public void TextDirectionTest(CellXf.TextDirectionValue value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(CellXf.DEFAULT_TEXT_DIRECTION, cellXf.TextDirection); // horizontal is default
            cellXf.TextDirection = value;
            Assert.Equal(value, cellXf.TextDirection);
            if (value == CellXf.TextDirectionValue.vertical)
            {
                Assert.Equal(255, cellXf.TextRotation);
            }
        }


        [Theory(DisplayName = "Test of the get and set function of the TextRotation property")]
        [InlineData(0)]
        [InlineData(33)]
        [InlineData(90)]
        [InlineData(-33)]
        [InlineData(-90)]
        public void TextRotationTest(int value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(0, cellXf.TextRotation); // 0 is default
            cellXf.TextRotation = value;
            Assert.Equal(value, cellXf.TextRotation);
        }

        [Theory(DisplayName = "Test of the failing get and set function of the TextRotation property on out-of-range values")]
        [InlineData(91)]
        [InlineData(-91)]
        [InlineData(-360)]
        [InlineData(360)]
        [InlineData(720)]
        public void TextRotationFailTest(int value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(0, cellXf.TextRotation); // 0 is default
            Assert.Throws<NanoXLSX.Exceptions.FormatException>(() => cellXf.TextRotation = value);
        }


        [Theory(DisplayName = "Test of the get and set function of the Align property")]
        [InlineData(CellXf.TextBreakValue.none)]
        [InlineData(CellXf.TextBreakValue.shrinkToFit)]
        [InlineData(CellXf.TextBreakValue.wrapText)]
        public void AlignTest(CellXf.TextBreakValue value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(CellXf.DEFAULT_ALIGNMENT, cellXf.Alignment); // none is default
            cellXf.Alignment = value;
            Assert.Equal(value, cellXf.Alignment);
        }

        [Theory(DisplayName = "Test of the get and set function of the Indent property")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(99)]
        public void IndentTest(int value)
        {
            CellXf cellXf = new CellXf();
            Assert.Equal(0, cellXf.Indent); // 0 is default
            cellXf.Indent = value;
            Assert.Equal(value, cellXf.Indent);
        }

        [Theory(DisplayName = "Test of the failing set function of the Indent property when an invalid value was passed")]
        [InlineData(-1)]
        [InlineData(-999)]
        public void IndentFailTest(int value)
        {
            Exception ex = Assert.Throws<StyleException>(() => exampleStyle.Indent = value);
            Assert.Equal(typeof(StyleException), ex.GetType());
        }

        [Fact(DisplayName = "Test of the Equals method")]
        public void EqualsTest()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            Assert.True(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Locked)")]
        public void EqualsTest2()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.Locked = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Hidden)")]
        public void EqualsTest2b()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.Hidden = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of HorizontalAlign)")]
        public void EqualsTest2c()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.HorizontalAlign = CellXf.HorizontalAlignValue.right;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of VerticalAlign)")]
        public void EqualsTest2d()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.VerticalAlign = CellXf.VerticalAlignValue.top;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of ForceApplyAlignment)")]
        public void EqualsTest2e()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.ForceApplyAlignment = false;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of TextDirection)")]
        public void EqualsTest2f()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.TextDirection = CellXf.TextDirectionValue.vertical;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of TextRotation)")]
        public void EqualsTest2g()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.TextRotation = 27;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Alignment)")]
        public void EqualsTest2h()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.Alignment = CellXf.TextBreakValue.none;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Fact(DisplayName = "Test of the Equals method (inequality of Indent)")]
        public void EqualsTest2i()
        {
            CellXf style2 = (CellXf)exampleStyle.Copy();
            style2.Indent = 77;
            Assert.False(exampleStyle.Equals(style2));
        }

        [Theory(DisplayName = "Test of the Equals method (inequality on null or different objects)")]
        [InlineData(null)]
        [InlineData("text")]
        [InlineData(true)]
        public void EqualsTest3(object obj)
        {
            Assert.False(exampleStyle.Equals(obj));
        }

        [Theory(DisplayName = "Test of the Equals method when the origin object is null or not of the same type")]
        [InlineData(null)]
        [InlineData(true)]
        [InlineData("origin")]
        public void EqualsTest5(object origin)
        {
            Assert.False(exampleStyle.Equals(origin));
        }

        [Fact(DisplayName = "Test of the GetHashCode method (equality of two identical objects)")]
        public void GetHashCodeTest()
        {
            CellXf copy = (CellXf)exampleStyle.Copy();
            copy.InternalID = 99;  // Should not influence
            Assert.Equal(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the GetHashCode method (inequality of two different objects)")]
        public void GetHashCodeTest2()
        {
            CellXf copy = (CellXf)exampleStyle.Copy();
            copy.Hidden = false;
            Assert.NotEqual(exampleStyle.GetHashCode(), copy.GetHashCode());
        }

        [Fact(DisplayName = "Test of the CompareTo method")]
        public void CompareToTest()
        {
            CellXf cellXf = new CellXf();
            CellXf other = new CellXf();
            cellXf.InternalID = null;
            other.InternalID = null;
            Assert.Equal(-1, cellXf.CompareTo(other));
            cellXf.InternalID = 5;
            Assert.Equal(1, cellXf.CompareTo(other));
            Assert.Equal(1, cellXf.CompareTo(null));
            other.InternalID = 5;
            Assert.Equal(0, cellXf.CompareTo(other));
            other.InternalID = 4;
            Assert.Equal(1, cellXf.CompareTo(other));
            other.InternalID = 6;
            Assert.Equal(-1, cellXf.CompareTo(other));
        }

        // For code coverage
        [Fact(DisplayName = "Test of the ToString function")]
        public void ToStringTest()
        {
            CellXf cellXf = new CellXf();
            string s1 = cellXf.ToString();
            cellXf.TextRotation = 12;
            Assert.NotEqual(s1, cellXf.ToString()); // An explicit value comparison is probably not sensible
        }

    }
}
