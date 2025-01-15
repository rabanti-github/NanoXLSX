using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NanoXLSX.Interfaces;
using Xunit;

namespace NanoXLSX.Core.Test.Misc
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
        [InlineData(LegacyPassword.PasswordType.WORKBOOK_PROTECTION)]
        [InlineData(LegacyPassword.PasswordType.WORKSHEET_PROTECTION)]
        public void ConstructorTest(LegacyPassword.PasswordType type)
        {
            LegacyPassword password = new LegacyPassword(type);
            Assert.NotNull(password);
            Assert.Equal(type, password.Type);
        }

        [Theory(DisplayName = "Test of the Type property (legacy)")]
        [InlineData(LegacyPassword.PasswordType.WORKBOOK_PROTECTION, LegacyPassword.PasswordType.WORKSHEET_PROTECTION)]
        [InlineData(LegacyPassword.PasswordType.WORKSHEET_PROTECTION, LegacyPassword.PasswordType.WORKBOOK_PROTECTION)]
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
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
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
        public void GetAndSetPasswordTest(string givenPassword, string expectedpassword)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            Assert.Null(password.GetPassword());
            password.SetPassword(givenPassword);
            Assert.Equal(expectedpassword, password.GetPassword());
        }

        [Theory(DisplayName = "Test of the UnsetPassword function (legacy)")]
        [InlineData("CEBA", true)]
        [InlineData("", false)]
        [InlineData("#@éü", true)]
        [InlineData(null, false)]
        [InlineData("0000", true)]
        public void UnsetPasswordTest(string plainText, bool expectedPasswordSet)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
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
        [InlineData("0000", true )]
        public void PasswordIsSetTest(string passwordHash, bool expectedPasswordSet)
        {
            LegacyPassword password = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            password.PasswordHash = passwordHash;
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
            LegacyPassword source = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
            source.SetPassword(plainText);
            LegacyPassword target = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            Assert.False(source.Equals(target));
            target.CopyFrom(source);
            Assert.True(source.Equals(target));
        }


        [Fact(DisplayName = "Test of the GetHashCode function (legacy)")]
        public void GetHashCodeTest()
        {
            LegacyPassword password1 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            password1.SetPassword("test");
            LegacyPassword password2 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            password2.SetPassword("test");
            LegacyPassword password3 = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
            password3.SetPassword(null);
            Assert.Equal(password1.GetHashCode(), password2.GetHashCode());
            Assert.NotEqual(password1.GetHashCode(), password3.GetHashCode());
        }

        [Fact(DisplayName = "Test of the Equals function (legacy)")]
        public void EqualsTest()
        {
            LegacyPassword password1 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            password1.SetPassword("test");
            LegacyPassword password2 = new LegacyPassword(LegacyPassword.PasswordType.WORKBOOK_PROTECTION);
            password2.SetPassword("test");
            LegacyPassword password3 = new LegacyPassword(LegacyPassword.PasswordType.WORKSHEET_PROTECTION);
            password3.SetPassword(null);
            Assert.True(password1.Equals(password2));
            Assert.False(password1.Equals(password3));
        }

    }
}
