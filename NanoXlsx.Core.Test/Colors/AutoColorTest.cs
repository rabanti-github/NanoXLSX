using NanoXLSX.Colors;
using Xunit;

namespace NanoXLSX.Core.Test.Colors
{
    public class AutoColorTest
    {
        [Fact(DisplayName = "Test of the getter of the StringValue property (dummy / for code completion)")]
        public void StringValueTest()
        {
            var color = new AutoColor();
            Assert.Null(color.StringValue); // Always null
        }
    }
}
