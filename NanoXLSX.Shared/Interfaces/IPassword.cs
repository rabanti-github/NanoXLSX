using System;
using System.Collections.Generic;
using System.Text;

namespace NanoXLSX.Interfaces
{
    public interface IPassword
    {
        string PasswordHash { get; set; }
        void SetPassword(string plainText);
        void UnsetPassword();
        string GetPassword();
        bool PasswordIsSet();
        void CopyFrom(IPassword passwordInstance);
    }
}
