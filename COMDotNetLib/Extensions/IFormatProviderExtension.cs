using System;

namespace DotNetLib.Extensions
{
    public static class IFormatProviderExtension
    {
        public static IFormatProvider Unwrap(this IFormatProvider provider)
        {
            if (provider is IWrappedObject unwrappedProvider)
            {
                return (IFormatProvider)unwrappedProvider.WrappedObject;
            }
            return provider; //If not a wrapped object return orignal object
        }

    }
}
