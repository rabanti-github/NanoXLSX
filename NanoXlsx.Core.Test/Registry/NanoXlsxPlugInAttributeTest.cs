using NanoXLSX.Registry.Attributes;
using Xunit;

namespace NanoXLSX.Test.Core.RegistryTest
{
    public class NanoXlsxPlugInAttributeTest
    {
        [Fact(DisplayName = "Default PlugInOrder should be 0")]
        public void DefaultPlugInOrderTest()
        {
            var attribute = new NanoXlsxPlugInAttribute();
            int actualOrder = attribute.PlugInOrder;
            Assert.Equal(0, actualOrder);
        }

        [Theory(DisplayName = "PlugInUUID Get/Set Test")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("UUID-123")]
        [InlineData("AnotherUniqueID")]
        public void PlugInUUIDGetSetTest(string expectedUUID)
        {
            var attribute = new NanoXlsxPlugInAttribute
            {
                PlugInUUID = expectedUUID
            };
            string actualUUID = attribute.PlugInUUID;
            Assert.Equal(expectedUUID, actualUUID);
        }

        [Theory(DisplayName = "PlugInOrder Get/Set Test")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(-1)]
        [InlineData(10)]
        public void PlugInOrderGetSetTest(int expectedOrder)
        {
            var attribute = new NanoXlsxPlugInAttribute
            {
                PlugInOrder = expectedOrder
            };
            int actualOrder = attribute.PlugInOrder;
            Assert.Equal(expectedOrder, actualOrder);
        }
    }
}
