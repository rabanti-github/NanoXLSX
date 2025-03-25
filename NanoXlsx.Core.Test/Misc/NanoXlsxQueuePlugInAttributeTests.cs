using NanoXLSX.Registry;
using Xunit;

namespace NanoXLSX.Test.Core.MiscTest
{
    public class NanoXlsxQueuePlugInAttributeTests
    {
        [Fact(DisplayName = "Default QueueUUID should be null")]
        public void DefaultQueueUUIDTest()
        {
            var attribute = new NanoXlsxQueuePlugInAttribute();
            Assert.Null(attribute.QueueUUID);
        }

        [Theory(DisplayName = "PlugInUUID Get/Set Test")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("UUID-123")]
        [InlineData("AnotherUniqueID")]
        public void PlugInUUIDGetSetTest(string expectedUUID)
        {
            var attribute = new NanoXlsxQueuePlugInAttribute();
            attribute.PlugInUUID = expectedUUID;
            Assert.Equal(expectedUUID, attribute.PlugInUUID);
        }

        [Theory(DisplayName = "QueueUUID Get/Set Test")]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("Queue-001")]
        [InlineData("AnotherQueue")]
        public void QueueUUIDGetSetTest(string expectedQueueUUID)
        {
            var attribute = new NanoXlsxQueuePlugInAttribute();
            attribute.QueueUUID = expectedQueueUUID;
            Assert.Equal(expectedQueueUUID, attribute.QueueUUID);
        }

        [Fact(DisplayName = "Default PlugInOrder should be 0")]
        public void DefaultPlugInOrderTest()
        {
            var attribute = new NanoXlsxQueuePlugInAttribute();
            Assert.Equal(0, attribute.PlugInOrder);
        }

        [Theory(DisplayName = "PlugInOrder Get/Set Test")]
        [InlineData(0)]
        [InlineData(1)]
        [InlineData(-1)]
        [InlineData(10)]
        public void PlugInOrderGetSetTest(int expectedOrder)
        {
            var attribute = new NanoXlsxQueuePlugInAttribute();
            attribute.PlugInOrder = expectedOrder;
            Assert.Equal(expectedOrder, attribute.PlugInOrder);
        }
    }
}
