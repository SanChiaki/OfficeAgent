using System;

namespace OfficeAgent.Core
{
    public sealed class AuthenticationRequiredException : InvalidOperationException
    {
        public AuthenticationRequiredException(string message)
            : base(message)
        {
        }
    }
}
