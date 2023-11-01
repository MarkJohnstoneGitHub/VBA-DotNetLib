// https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.-ctor?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("71BCF1BB-F1B1-4494-9E88-4C5E0E614813")]
    [Description("Manages a compact array of bit values, which are represented as Booleans, where true indicates that the bit is on (1) and false indicates the bit is off (0).")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBitArray
    {
        int Count 
        {
            [Description("Gets the number of elements contained in the BitArray.")]
            get;
        }

        bool IsReadOnly 
        {
            [Description("Gets a value indicating whether the BitArray is read-only.")]
            get;
        }

        bool IsSynchronized 
        {
            [Description("Gets a value indicating whether access to the BitArray is synchronized (thread safe).")]
            get;
        }

        bool this[int index] 
        {
            [Description("Gets or sets the value of the bit at a specific position in the BitArray.")]
            get;
            [Description("Gets or sets the value of the bit at a specific position in the BitArray.")]
            set;
        }

        int Length 
        {
            [Description("Gets or sets the number of elements in the BitArray.")]
            get;
            [Description("Gets or sets the number of elements in the BitArray.")]
            set;
        }

        object SyncRoot 
        {
            [Description("Gets an object that can be used to synchronize access to the BitArray.")]
            get;
        }

        // Methods

        [Description("Performs the bitwise AND operation between the elements of the current BitArray object and the corresponding elements in the specified array. The current BitArray object will be modified to store the result of the bitwise AND operation.")]
        BitArray And(BitArray value);

        [Description("Copies the entire BitArray to a compatible one-dimensional Array, starting at the specified index of the target array.")]
        void CopyTo(Array array, int index);

        [Description("")]
        bool Get(int index);

        [Description("Returns an enumerator that iterates through the BitArray.")]
        GCollections.IEnumerator GetEnumerator();

        [Description("Inverts all the bit values in the current BitArray, so that elements set to true are changed to false, and elements set to false are changed to true.")] 
        BitArray Not();

        [Description("Performs the bitwise OR operation between the elements of the current BitArray object and the corresponding elements in the specified array. The current BitArray object will be modified to store the result of the bitwise OR operation.")]
        BitArray Or(BitArray value);

        [Description("Sets the bit at a specific position in the BitArray to the specified value.")]
        void Set(int index, bool value);

        [Description("Sets all bits in the BitArray to the specified value.")]
        void SetAll(bool value);

        [Description("Performs the bitwise exclusive OR operation between the elements of the current BitArray object against the corresponding elements in the specified array. The current BitArray object will be modified to store the result of the bitwise exclusive OR operation.")]
        BitArray Xor(BitArray value);

        [Description("Serves as the default hash function.\r\n\r\n(Inherited from Object)")]
        int GetHashCode();

        [Description("Gets the Type of the current instance.\r\n\r\n(Inherited from Object)")]
        Type GetType();

        [Description("Returns a string that represents the current object.\r\n\r\n(Inherited from Object)")]
        string ToString();

        //[Description("")]
    }
}
