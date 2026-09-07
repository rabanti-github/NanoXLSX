using Xunit;

using System;
using System.Reflection;
using System.Security;
using NanoXLSX.Interfaces;

namespace NanoXLSX.Test.Core.MiscTest
{
    public class LegacyPasswordTest
    {
        [Theory(DisplayName = "Test of the GeneratePasswordHash function (legacy)")]
        [InlineData("x", "CEBA")]
        [InlineData("Test@1-2,3!", "F767")]
        [InlineData(" ", "CE0A")]
        [InlineData("", "")]
        [InlineData(null, "")]
        public void GeneratePasswordHashTest(string givenVPassword, string expectedHash)
        {
            string hash = LegacyPassword.GenerateLegacyPasswordHash(givenVPassword);
            Assert.Equal(expectedHash, hash);
        }

        [Theory(DisplayName = "Test of the LegacyPassword constructor with arguments (legacy)")]
        [InlineData(LegacyPassword.PasswordType.WorkbookProtection)]
        [InlineData(LegacyPassword.PasswordType.WorksheetProtection)]
        public void ConstructorTest(LegacyPassword.PasswordType type)
        {
            LegacyPassword password = new LegacyPassword(type);
            Assert.NotNull(password);
            Assert.Equal(type, password.Type);
        }

        [Theory(DisplayName = "Test of the Type property (legacy)")]
        [InlineData(LegacyPassword.PasswordType.WorkbookProtection, LegacyPassword.PasswordType.WorksheetProtection)]
        [InlineData(LegacyPassword.PasswordType.WorksheetProtection, LegacyPassword.PasswordType.WorkbookProtection)]
        public void PasswordTypeTest(LegacyPassword.PasswordType initialType, LegacyPassword.PasswordType type)
        {
            LegacyPassword password = new LegacyPassword(initialType);
            Assert.Equal(initialType, password.Type);
            password.Type = type;
            Assert.Equal(type, password.Type);
        }

        [Theory(DisplayName = "Test of the PasswordHash property (legacy)")]
        [InlineData("CEBA")]
        [InlineData("")]
        [InlineData(null)]
        [InlineData("0000")]
        public void PasswordHashTest(string passwordHash)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            Assert.Null(password.PasswordHash);
            password.PasswordHash = passwordHash;
            Assert.Equal(passwordHash, password.PasswordHash);
        }

        [Theory(DisplayName = "Test of the GetPassword and SetPassword functions (legacy)")]
        [InlineData("test", "test")]
        [InlineData("0123", "0123")]
        [InlineData("#@éü", "#@éü")]
        [InlineData(" ", " ")]
        [InlineData(null, null)]
        [InlineData("", null)]
        public void GetAndSetPasswordTest(string givenPassword, string expectedPassword)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            Assert.Null(password.GetPassword());
            password.SetPassword(givenPassword);
            Assert.Equal(expectedPassword, password.GetPassword());
        }

        [Theory(DisplayName = "Test of the UnsetPassword function (legacy)")]
        [InlineData("CEBA", true)]
        [InlineData("", false)]
        [InlineData("#@éü", true)]
        [InlineData(null, false)]
        [InlineData("0000", true)]
        public void UnsetPasswordTest(string plainText, bool expectedPasswordSet)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            Assert.Null(password.PasswordHash);
            password.SetPassword(plainText);
            if (expectedPasswordSet)
            {
                Assert.True(password.PasswordIsSet());
                Assert.Equal(plainText, password.GetPassword());
                Assert.NotNull(password.PasswordHash);
            }
            else
            {
                Assert.False(password.PasswordIsSet());
                Assert.Null(password.GetPassword());
                Assert.Null(password.PasswordHash);
            }

            password.UnsetPassword();
            Assert.False(password.PasswordIsSet());
            Assert.Null(password.GetPassword());
            Assert.Null(password.PasswordHash);
        }

        [Theory(DisplayName = "Test of the PasswordIsSet function (legacy)")]
        [InlineData("CEBA", true)]
        [InlineData("", false)]
        [InlineData("#@éü", true)]
        [InlineData(null, false)]
        [InlineData("0000", true)]
        public void PasswordIsSetTest(string passwordHash, bool expectedPasswordSet)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection)
            {
                PasswordHash = passwordHash
            };
            Assert.Equal(expectedPasswordSet, password.PasswordIsSet());
        }

        [Theory(DisplayName = "Test of the CopyFromTest function (legacy)")]
        [InlineData("CEBA")]
        [InlineData("")]
        [InlineData("#@éü")]
        [InlineData(null)]
        [InlineData("0000")]
        public void CopyFromTest(string plainText)
        {
            LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WorksheetProtection);
            source.SetPassword(plainText);
            LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            Assert.False(source.Equals(target));
            target.CopyFrom(source);
            Assert.True(source.Equals(target));
        }

        [Fact(DisplayName = "Test of the CopyFrom function with a null source (legacy)")]
        public void CopyFromNullTest()
        {
            LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            target.SetPassword("original");

            ArgumentNullException exception = Assert.Throws<ArgumentNullException>(() => target.CopyFrom(null));

            Assert.Equal("passwordInstance", exception.ParamName);
            Assert.Equal("original", target.GetPassword());
            Assert.Equal(LegacyPassword.PasswordType.WorkbookProtection, target.Type);
        }

        [Fact(DisplayName = "Test of independent password storage after CopyFrom (legacy)")]
        public void CopyFromCreatesIndependentPasswordStorageTest()
        {
            LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WorksheetProtection);
            source.SetPassword("source");
            LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            target.SetPassword("target");
            SecureString previousTargetPassword = GetSecurePassword(target);

            target.CopyFrom(source);
            source.UnsetPassword();

            Assert.Throws<ObjectDisposedException>(() => previousTargetPassword.AppendChar('x'));
            Assert.Equal("source", target.GetPassword());
            Assert.Equal(LegacyPassword.PasswordType.WorksheetProtection, target.Type);
        }

        [Fact(DisplayName = "Test of CopyFrom preparation failure (legacy)")]
        public void CopyFromPreparationFailureTest()
        {
            LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            target.SetPassword("original");
            IPassword source = new ThrowingPassword();

            Assert.Throws<InvalidOperationException>(() => target.CopyFrom(source));

            Assert.Equal("original", target.GetPassword());
            Assert.Equal(LegacyPassword.PasswordType.WorkbookProtection, target.Type);
        }

        [Fact(DisplayName = "Test of CopyFrom with another IPassword implementation (legacy)")]
        public void CopyFromInterfaceTest()
        {
            IPassword source = new TestPassword("source", "F462");
            LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);

            target.CopyFrom(source);

            Assert.Equal("source", target.GetPassword());
            Assert.Equal("F462", target.PasswordHash);
            Assert.Equal(LegacyPassword.PasswordType.WorkbookProtection, target.Type);
        }

        [Fact(DisplayName = "Test of the Dispose function (legacy)")]
        public void DisposeTest()
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorksheetProtection);
            password.SetPassword("test");
            SecureString securePassword = GetSecurePassword(password);

            password.Dispose();
            password.Dispose();

            Assert.Throws<ObjectDisposedException>(() => securePassword.AppendChar('x'));
            Assert.Null(password.GetPassword());
            Assert.Null(password.PasswordHash);
            Assert.False(password.PasswordIsSet());
            Assert.Equal(LegacyPassword.PasswordType.WorksheetProtection, password.Type);
        }

        [Fact(DisplayName = "Test of the protected Dispose function (legacy)")]
        public void ProtectedDisposeTest()
        {
            DisposableLegacyPassword password = new DisposableLegacyPassword(
                LegacyPassword.PasswordType.WorkbookProtection);
            password.SetPassword("test");

            password.InvokeDispose(false);

            Assert.Equal("test", password.GetPassword());
            Assert.True(password.PasswordIsSet());

            password.InvokeDispose(true);

            Assert.Null(password.GetPassword());
            Assert.False(password.PasswordIsSet());
        }

        [Fact(DisplayName = "Test of reuse after Dispose (legacy)")]
        public void ReuseAfterDisposeTest()
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            password.SetPassword("first");
            password.Dispose();

            password.SetPassword("second");

            Assert.Equal("second", password.GetPassword());
            Assert.True(password.PasswordIsSet());
            password.Dispose();
        }

        [Fact(DisplayName = "Test of SecureString cleanup after password preparation failure (legacy)")]
        public void SecureStringPreparationFailureTest()
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            password.SetPassword("original");
            string oversizedPassword = new string('x', 65537);

            Assert.Throws<ArgumentOutOfRangeException>(() => password.SetPassword(oversizedPassword));

            Assert.Equal("original", password.GetPassword());
            Assert.True(password.PasswordIsSet());
        }


        [Fact(DisplayName = "Test of the GetHashCode function (legacy)")]
        public void GetHashCodeTest()
        {
            LegacyPassword password1 = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            password1.SetPassword("test");
            LegacyPassword password2 = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            password2.SetPassword("test");
            LegacyPassword password3 = new LegacyPassword(LegacyPassword.PasswordType.WorksheetProtection);
            password3.SetPassword(null);
            Assert.Equal(password1.GetHashCode(), password2.GetHashCode());
            Assert.NotEqual(password1.GetHashCode(), password3.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Equals function (legacy)")]
        public void EqualsTest()
        {
            LegacyPassword password1 = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            password1.SetPassword("test");
            LegacyPassword password2 = new LegacyPassword(LegacyPassword.PasswordType.WorkbookProtection);
            password2.SetPassword("test");
            LegacyPassword password3 = new LegacyPassword(LegacyPassword.PasswordType.WorksheetProtection);
            password3.SetPassword(null);
            Assert.True(password1.Equals(password2));
            Assert.False(password1.Equals(password3));
        }

        private static SecureString GetSecurePassword(LegacyPassword password)
        {
            FieldInfo field = typeof(LegacyPassword).GetField(
                "password",
                BindingFlags.Instance | BindingFlags.NonPublic);
            return (SecureString)field.GetValue(password);
        }

        private sealed class DisposableLegacyPassword : LegacyPassword
        {
            public DisposableLegacyPassword(PasswordType type) : base(type)
            {
            }

            public void InvokeDispose(bool disposing)
            {
                base.Dispose(disposing);
            }
        }

        private sealed class TestPassword : IPassword
        {
            private string plainText;

            public TestPassword(string plainText, string passwordHash)
            {
                this.plainText = plainText;
                PasswordHash = passwordHash;
            }

            public string PasswordHash { get; set; }

            public void SetPassword(string value)
            {
                plainText = value;
            }

            public void UnsetPassword()
            {
                plainText = null;
                PasswordHash = null;
            }

            public string GetPassword()
            {
                return plainText;
            }

            public bool PasswordIsSet()
            {
                return !string.IsNullOrEmpty(PasswordHash);
            }

            public void CopyFrom(IPassword passwordInstance)
            {
                plainText = passwordInstance.GetPassword();
                PasswordHash = passwordInstance.PasswordHash;
            }
        }

        private sealed class ThrowingPassword : IPassword
        {
            public string PasswordHash { get; set; } = "hash";

            public void SetPassword(string plainText)
            {
                throw new NotImplementedException();
            }

            public void UnsetPassword()
            {
                throw new NotImplementedException();
            }

            public string GetPassword()
            {
                throw new InvalidOperationException("Password retrieval failed");
            }

            public bool PasswordIsSet()
            {
                throw new NotImplementedException();
            }

            public void CopyFrom(IPassword passwordInstance)
            {
                throw new NotImplementedException();
            }
        }

    }
}
