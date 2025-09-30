﻿using NanoXLSX.Exceptions;
using System;
using Xunit;
using FormatException = NanoXLSX.Exceptions.FormatException;

namespace NanoXLSX_Test.Misc
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
        }

    }
}
