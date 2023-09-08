using NanoXLSX.Shared.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;
using Xunit;
using FormatException = NanoXLSX.Shared.Exceptions.FormatException;

namespace NanoXLSX.Shared_Test.Misc
{
    public class ExceptionTest
    {
        // For code coverage
        [Fact(DisplayName = "Test of the FormatException (summary)")]
        public void FormatExceptionTest()
        {
            FormatException exception = new FormatException();
            Assert.NotEmpty(exception.Message); // Gets a generated message my the base class
            Assert.Null(exception.InnerException);

            exception = new FormatException("test");
            Assert.Equal("test", exception.Message);
            Assert.Null(exception.InnerException);

            AssertExceptionSerialization<FormatException>(exception);

            ArgumentException inner = new ArgumentException("inner message");
            exception = new FormatException("test", inner);
            Assert.Equal("test", exception.Message);
            Assert.NotNull(exception.InnerException);
            Assert.Equal(typeof(ArgumentException), exception.InnerException.GetType());
            Assert.Equal("inner message", exception.InnerException.Message);
        }

        [Fact(DisplayName = "Test of the  IOExceptio (summary)")]
        public void IOExceptionTest()
        {
            IOException exception = new IOException();
            Assert.NotEmpty(exception.Message); // Gets a generated message my the base class
            Assert.Null(exception.InnerException);

            exception = new IOException("test");
            Assert.Equal("test", exception.Message);
            Assert.Null(exception.InnerException);

            AssertExceptionSerialization<IOException>(exception);

            ArgumentException inner = new ArgumentException("inner message");
            exception = new IOException("test", inner);
            Assert.Equal("test", exception.Message);
            Assert.NotNull(exception.InnerException);
            Assert.Equal(typeof(ArgumentException), exception.InnerException.GetType());
            Assert.Equal("inner message", exception.InnerException.Message);
        }

        [Fact(DisplayName = "Test of the RangeException (summary)")]
        public void RangeExceptionTest()
        {
            RangeException exception = new RangeException();
            Assert.NotEmpty(exception.Message); // Gets a generated message my the base class
            Assert.Null(exception.InnerException);

            exception = new RangeException("test");
            Assert.Equal("test", exception.Message);
            Assert.Null(exception.InnerException);

            AssertExceptionSerialization<RangeException>(exception);
        }

        [Fact(DisplayName = "Test of the  StyleException (summary)")]
        public void StyleExceptionTest()
        {
            StyleException exception = new StyleException();
            Assert.NotEmpty(exception.Message); // Gets a generated message my the base class
            Assert.Null(exception.InnerException);

            exception = new StyleException("test");
            Assert.Equal("test", exception.Message);
            Assert.Null(exception.InnerException);

            AssertExceptionSerialization<StyleException>(exception);

            ArgumentException inner = new ArgumentException("inner message");
            exception = new StyleException("test", inner);
            Assert.Equal("test", exception.Message);
            Assert.NotNull(exception.InnerException);
            Assert.Equal(typeof(ArgumentException), exception.InnerException.GetType());
            Assert.Equal("inner message", exception.InnerException.Message);
        }

        [Fact(DisplayName = "Test of the WorksheetException (summary)")]
        public void WorksheetExceptionTest()
        {
            WorksheetException exception = new WorksheetException();
            Assert.NotEmpty(exception.Message); // Gets a generated message my the base class
            Assert.Null(exception.InnerException);

            exception = new WorksheetException("test");
            Assert.Equal("test", exception.Message);
            Assert.Null(exception.InnerException);

            AssertExceptionSerialization<WorksheetException>(exception);
        }

        public static void AssertExceptionSerialization<TException>(TException originalException) where TException : Exception
        {
            BinaryFormatter formatter = new BinaryFormatter();
            TException deserializedException;

            using (var stream = new System.IO.MemoryStream())
            {
                formatter.Serialize(stream, originalException);

                stream.Seek(0, System.IO.SeekOrigin.Begin);
                deserializedException = (TException)formatter.Deserialize(stream);
            }
            Assert.Equal(originalException.Message, deserializedException.Message);
        }

    }
}
