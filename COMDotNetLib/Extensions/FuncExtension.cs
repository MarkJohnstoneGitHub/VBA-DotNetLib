// https://stackoverflow.com/questions/731249/how-to-convert-funct-bool-to-predicatet

using GSystem = global::System;
using System;
using System.Runtime.InteropServices;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public static class FuncExtension
    {
        public static GSystem.Predicate<T> ToPredicate<T>(this Func<T, bool> f)
        {
            return x => f(x);
        }
    }
}
