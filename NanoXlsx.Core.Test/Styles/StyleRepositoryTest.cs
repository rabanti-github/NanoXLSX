using NanoXLSX.Styles;
using NanoXLSX.Test.Core.Utils;
using System;
using System.Collections.Generic;
using Xunit;

namespace NanoXLSX.Test.Core.StyleTest
{
    // Ensure that these tests are executed sequentially, since static repository methods may be called 
    [Collection(nameof(SequentialCollection))]
    public class StyleRepositoryTest
    {
        public StyleRepositoryTest()
        {
            StyleRepository.Instance.FlushStyles();
        }

        [Fact(DisplayName = "Test of the AddStyle method")]
        public void AddStyleTest()
        {
            StyleRepository repository = StyleRepository.Instance;
            Assert.Empty(repository.ManagedStyles);
            Style style = new Style();
            style.CurrentFont.Name = "Arial";
            Style result = repository.AddStyle(style);
            Assert.Single(repository.ManagedStyles);
            Assert.Equal(style.GetHashCode(), result.GetHashCode());
            Assert.Equal(style.GetHashCode(), repository.ManagedStyles[style.GetHashCode()].GetHashCode());
        }

        [Fact(DisplayName = "Test of the AddStyle method on a null object")]
        public void AddStyleTest2()
        {
            StyleRepository repository = StyleRepository.Instance;
            Assert.Empty(repository.ManagedStyles);
            Style result = repository.AddStyle(null);
            Assert.Empty(repository.ManagedStyles);
            Assert.Null(result);
        }

        [Fact(DisplayName = "Test of the Flush method")]
        public void FlushTest()
        {
            StyleRepository repository = StyleRepository.Instance;
            Assert.Empty(repository.ManagedStyles);
            Style style = new Style();
            style.CurrentFont.Name = "Arial";
            repository.AddStyle(style);
            Assert.Single(repository.ManagedStyles);
            repository.FlushStyles();
            Assert.Empty(repository.ManagedStyles);
        }

        [Fact(DisplayName = "Test that the obsolete Styles property cannot mutate the repository")]
        public void StylesSnapshotTest()
        {
            StyleRepository repository = StyleRepository.Instance;
            Style style = new Style();
            repository.AddStyle(style);

            Dictionary<int, Style> styles = repository.Styles;
            styles.Clear();
            Assert.Single(repository.ManagedStyles);
        }

        [Fact(DisplayName = "Test that the ManagedStyles property is read-only")]
        public void ManagedStylesReadOnlyTest()
        {
            StyleRepository repository = StyleRepository.Instance;
            Style style = new Style();
            repository.AddStyle(style);

            IDictionary<int, Style> styles = Assert.IsAssignableFrom<IDictionary<int, Style>>(repository.ManagedStyles);
            Assert.Throws<NotSupportedException>(() => styles.Clear());
            Assert.Single(repository.ManagedStyles);
        }

    }
}
