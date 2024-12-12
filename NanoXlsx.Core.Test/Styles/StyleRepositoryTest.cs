using NanoXLSX.Styles;
using Xunit;

namespace NanoXLSX.Test.Styles
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class StyleRepositoryTest
    {
        public StyleRepositoryTest()
        {
            StyleRepository.Instance.Styles.Clear();
        }

        [Fact(DisplayName = "Test of the AddStyle method")]
        public void AddStyleTest()
        {
            StyleRepository repository = StyleRepository.Instance;
            Assert.Empty(repository.Styles);
            Style style = new Style();
            style.CurrentFont.Name = "Arial";
            Style result = repository.AddStyle(style);
            Assert.Single(repository.Styles);
            Assert.Equal(style.GetHashCode(), result.GetHashCode());
            Assert.Equal(style.GetHashCode(), repository.Styles[style.GetHashCode()].GetHashCode());
        }

        [Fact(DisplayName = "Test of the AddStyle method on a null object")]
        public void AddStyleTest2()
        {
            StyleRepository repository = StyleRepository.Instance;
            Assert.Empty(repository.Styles);
            Style result = repository.AddStyle(null);
            Assert.Empty(repository.Styles);
            Assert.Null(result);
        }

        [Fact(DisplayName = "Test of the Flush method")]
        public void FlushTest()
        {
            StyleRepository repository = StyleRepository.Instance;
            Assert.Empty(repository.Styles);
            Style style = new Style();
            style.CurrentFont.Name = "Arial";
            repository.AddStyle(style);
            Assert.Single(repository.Styles);
            repository.FlushStyles();
            Assert.Empty(repository.Styles);
        }

    }
}
