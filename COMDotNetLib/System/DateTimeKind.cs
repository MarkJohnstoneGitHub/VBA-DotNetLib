using System.Runtime.InteropServices;

// https://stackoverflow.com/questions/11647647/how-to-expose-an-enum-defined-in-a-com-library-via-interop-as-the-return-type-of

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("3D9FEAA7-82DF-4906-A092-2E76E412B966")]

    public enum DateTimeKind
    {
        Unspecified = 0,
        Utc = 1,
        Local = 2,
    }
}
