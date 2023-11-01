// https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.-ctor?view=netframework-4.8.1

using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Manages a compact array of bit values, which are represented as Booleans, where true indicates that the bit is on (1) and false indicates the bit is off (0).")]
    [Guid("A97BB12C-E2E2-40A0-BD1F-185C3887758E")]
    [ProgId("DotNetLib.System.Collections.BitArraySingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBitArraySingleton))]
    public class BitArraySingleton : IBitArraySingleton
    {
        public BitArraySingleton() { }

        public BitArray Create(int length)
        {
            return new BitArray(length);
        }

        public BitArray Create(int length, bool defaultValue)
        {
            return new BitArray(length, defaultValue);
        }

        public BitArray Create([In] ref byte[] bytes)
        {
            return new BitArray(bytes);
        }

        public BitArray Create([In] ref bool[] values)
        {
            return new BitArray(values);
        }

        public BitArray Create([In] ref int[] values)
        {
            return new BitArray(values);
        }

        public BitArray Create(BitArray bits)
        {

            return new BitArray(bits);
        }

    }
}
