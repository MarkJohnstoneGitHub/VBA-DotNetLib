// https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.-ctor?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("6873E657-4A8A-4A3F-9074-68B0EDE37F44")]
    [Description("Manages a compact array of bit values, which are represented as Booleans, where true indicates that the bit is on (1) and false indicates the bit is off (0).")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]

    public interface IBitArraySingleton
    {
        [Description("Initializes a new instance of the BitArray class that can hold the specified number of bit values, which are initially set to false.")]
        BitArray Create(int length);

        [Description("Initializes a new instance of the BitArray class that can hold the specified number of bit values, which are initially set to the specified value.")]
        BitArray Create(int length, bool defaultValue);

        [Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified array of bytes.")]
        BitArray Create([In] ref byte[] bytes);

        [Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified array of Booleans.")]
        BitArray Create([In] ref bool[] values);

        [Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified array of 32-bit integers.")]
        BitArray Create([In] ref int[] values);

        [Description("Initializes a new instance of the BitArray class that contains bit values copied from the specified BitArray.")]
        BitArray Create(BitArray bits);
    }
}
