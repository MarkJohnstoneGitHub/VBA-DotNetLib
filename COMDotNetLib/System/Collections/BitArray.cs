// https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray.-ctor?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Manages a compact array of bit values, which are represented as Booleans, where true indicates that the bit is on (1) and false indicates the bit is off (0).")]
    [Guid("C01910B5-D811-48B8-ABF1-E4B08A3D2FAB")]
    [ProgId("DotNetLib.System.Collections.BitArray")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IBitArray))]
    public class BitArray  : GCollections.ICollection, GCollections.IEnumerable, GSystem.ICloneable, IBitArray
    {
        private GCollections.BitArray _bitArray;

        // Constructors
        public BitArray(System.Collections.BitArray bits)
        {
            _bitArray = new GCollections.BitArray(bits.WrappedBitArray);
        }

        public BitArray(int length)
        {
            _bitArray = new GCollections.BitArray(length);
        }

        public BitArray(bool[] values)
        {
            _bitArray = new GCollections.BitArray(values);
        }

        public BitArray(int[] values)
        {
            _bitArray = new GCollections.BitArray(values);
        }

        public BitArray(byte[] bytes)
        {
            _bitArray = new GCollections.BitArray(bytes);
        }

        public BitArray(int length, bool defaultValue)
        {
            _bitArray = new GCollections.BitArray(length, defaultValue);
        }

        //Properties

        internal GCollections.BitArray WrappedBitArray
        {
            get { return _bitArray; }
            set { _bitArray = value; }
        }

        public int Count => _bitArray.Count;

        public bool IsSynchronized => _bitArray.IsSynchronized;

        public bool IsReadOnly => _bitArray.IsReadOnly;

        public bool this[int index] 
        {
            get => _bitArray[index];
            set => _bitArray[index] = value;
        }

        public int Length {
            get => _bitArray.Length;
            set => _bitArray.Length = value;
        }

        public object SyncRoot => _bitArray.SyncRoot;


        // Methods

        public BitArray And(BitArray value)
        {
            _bitArray.And(value.WrappedBitArray);
            return this;
        }

        public void CopyTo(Array array, int index)
        {
            _bitArray.CopyTo(array.WrappedArray, index);
        }

        public BitArray Not()
        {
            _bitArray.Not();
            return this;
        }

        public BitArray Or(BitArray value)
        {
            _bitArray.Or(value.WrappedBitArray);
            return this;
        }

        public void Set(int index, bool value)
        {
            _bitArray.Set(index,value);
        }

        public void SetAll(bool value)
        { 
            _bitArray.SetAll(value);
        }

        //Todo: check implementaation suspect cloning twice?
        public object Clone()
        {
            return new BitArray((BitArray)_bitArray.Clone());
        }

        public void CopyTo(global::System.Array array, int index)
        {
            _bitArray.CopyTo(array, index);
        }

        public bool Get(int index)
        { 
            return _bitArray.Get(index);
        }

        public BitArray Xor(BitArray value)
        {
            _bitArray.Xor(value.WrappedBitArray);
            return this;
        }

        public GCollections.IEnumerator GetEnumerator()
        {
            return _bitArray.GetEnumerator();
        }

        Type IBitArray.GetType()
        {
            return new Type(((object)this).GetType());
        }


    }
}
