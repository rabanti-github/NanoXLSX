using NanoXLSX.Styles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace NanoXLSX_Test.Styles
{
    public class AppendAttributeTest
    {

        [Theory(DisplayName = "Test of the AppendAttribute, applied to a dummy class")]
        [InlineData("AppendProperty", true, false, false)]
        [InlineData("AppendPropertyNonIgnore", true, false, false)]
        [InlineData("IgnoreProperty", true, true, false)]
        [InlineData("NestedProperty", true, false, true)]
        [InlineData("NonNestedProperty", true, false, false)]
        [InlineData("UndefinedProperty", false, false, false)]
        public void AppendAttributeTest1(String propertyName, bool expectedAttribute, bool expectedIgnore, bool expectedNested)
        {

            PropertyInfo[] propertiesInfo = typeof(DummyClass).GetProperties();
            bool propertyFound = false;
            foreach(PropertyInfo info in propertiesInfo)
            {
                if (info.Name == propertyName)
                {
                    object[] attributes = info.GetCustomAttributes(true);
                    bool attributeFound = false;
                    foreach(object attribute in attributes)
                    {
                        AppendAttribute appendAttribute = attribute as AppendAttribute;
                        if (appendAttribute != null)
                        {
                            Assert.True(expectedAttribute);
                            Assert.Equal(expectedIgnore, appendAttribute.Ignore);
                            Assert.Equal(expectedNested, appendAttribute.NestedProperty);
                            attributeFound = true;
                        }
                    }
                    propertyFound = true;
                    if (expectedAttribute)
                    {
                        Assert.True(attributeFound);
                    }
                }
            }
            if (expectedAttribute)
            {
                Assert.True(propertyFound);
            }
        }

        private class DummyClass
        {
            [Append]
            public int AppendProperty { get; set; }
            [Append (Ignore = false)]
            public int AppendPropertyNonIgnore { get; set; }
            [Append(Ignore = true)]
            public int IgnoreProperty { get; set; }
            [Append(NestedProperty = true)]
            public int NestedProperty { get; set; }
            [Append(NestedProperty = false)]
            public int NonNestedProperty { get; set; }
            public int UndefinedProperty { get; set; }
        }

    }
}
