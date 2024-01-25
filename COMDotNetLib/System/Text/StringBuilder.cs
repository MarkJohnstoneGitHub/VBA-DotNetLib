// https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1

using GText = global::System.Text;
using System;
using System.Runtime.Serialization;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.ComponentModel;

namespace DotNetLib.System.Text
{
    [ComVisible(true)]
    [Description("Represents a mutable string of characters. This class cannot be inherited.")]
    [Guid("B3D450A7-CFFC-48C8-A157-8755032BBF3F")]
    [ProgId("DotNetLib.System.Text.StringBuilder")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStringBuilder))]
    public class StringBuilder : ISerializable, IWrappedObject, IStringBuilder
    {
        private GText.StringBuilder _sb;

        public StringBuilder()
        {
            _sb = new GText.StringBuilder();  
        }

        public StringBuilder(int capacity) 
        {
            _sb = new GText.StringBuilder(capacity);
        }

        public StringBuilder(string value) 
        { 
            _sb = new GText.StringBuilder(value);
        }
        public StringBuilder(string value, int capacity) 
        { 
            _sb = new GText.StringBuilder(value, capacity);
        }

        public StringBuilder(int capacity, int maxCapacity)
        { 
            _sb = new GText.StringBuilder(capacity, maxCapacity);
        }

        public StringBuilder(string value, int startIndex, int length, int capacity)
        { 
            _sb = new GText.StringBuilder(value,startIndex, length, capacity);
        }

        internal StringBuilder(GText.StringBuilder stringBuilder)
        {
            _sb = stringBuilder;
        }

        // Properties

        public GText.StringBuilder WrappedStringBuilder
        {
            get { return _sb; }
            set
            {
                _sb = value;
            }
        }

        public object WrappedObject => _sb;

        public int Capacity 
        { 
            get => _sb.Capacity;
            set => _sb.Capacity = value;
        }
        public int Length 
        {
            get => _sb.Length;
            set => _sb.Length = value;
        }

        public int MaxCapacity => _sb.MaxCapacity;

        public string this[int index]
        {
            get => _sb[index].ToString();

            set => _sb[index] = value[0];
        }


        // Methods

        public StringBuilder Append(string value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(string value, int startIndex, int count)
        {
            _sb.Append(value, startIndex, count);
            return this;
        }

        public StringBuilder Append(StringBuilder value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(bool value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(byte value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(short value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(int value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(long value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(float value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append(double value)
        {
            _sb.Append(value);
            return this;
        }

        //public StringBuilder Append(decimal value)
        //{
        //    _sb.Append(value);
        //    return this;
        //}

        public StringBuilder Append(object value)
        {
            _sb.Append(value.Unwrap());
            return this;
        }

        public StringBuilder AppendFormat(string format, [In] ref object[] args)
        {
            _sb.AppendFormat(format, args);
            return this;
        }

        public StringBuilder AppendFormat(string format, object arg0)
        {
            _sb.AppendFormat(format, arg0);
            return this;
        }

        public StringBuilder AppendFormat(string format, object arg0, object arg1)
        {
            _sb.AppendFormat(format, arg0, arg1);
            return this;
        }

        public StringBuilder AppendFormat(string format, object arg0, object arg1, object arg2)
        {
            _sb.AppendFormat(format, arg0, arg1, arg2);
            return this;
        }

        public StringBuilder AppendFormat(IFormatProvider provider, string pFormat, [In] ref object[] args)
        {
            _sb.AppendFormat(provider, pFormat, args);
            return this;
        }

        public StringBuilder AppendFormat(IFormatProvider provider, string format, object arg0)
        {
            _sb.AppendFormat(provider, format, arg0);
            return this;
        }

        public StringBuilder AppendFormat(IFormatProvider provider, string format, object arg0, object arg1)
        {
            _sb.AppendFormat(provider, format, arg0, arg1);
            return this;
        }

        public StringBuilder AppendFormat(IFormatProvider provider, string format, object arg0, object arg1, object arg2)
        {
            _sb.AppendFormat(provider, format, arg0, arg1, arg2);
            return this;
        }


        public StringBuilder AppendLine()
        {
            _sb.AppendLine();
            return this;
        }

        public StringBuilder AppendLine(string value)
        {
            _sb.AppendLine(value);
            return this;
        }

        public StringBuilder Clear()
        {
            _sb.Clear();
            return this;
        }

        public int EnsureCapacity(int capacity)
        {
            return _sb.EnsureCapacity(capacity);
        }

        public bool Equals(StringBuilder sb)
        {
            return _sb.Equals(sb.WrappedStringBuilder);
        }


        public StringBuilder Insert(int index, string value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, string value, int count)
        {
            _sb.Insert(index, value, count);
            return this;
        }

        public StringBuilder Insert(int index, bool value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, byte value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, short value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, int value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, long value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, float value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, double value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, decimal value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, object value)
        {
            _sb.Insert(index, value.Unwrap());
            return this;
        }

        public StringBuilder Remove(int startIndex, int length)
        {
            _sb.Remove(startIndex, length);
            return this;
        }

        public StringBuilder Replace(string oldValue, string newValue)
        {
            _sb.Replace(oldValue, newValue);
            return this;
        }

        public StringBuilder Replace(string oldValue, string newValue, int startIndex, int count)
        {
            _sb.Replace(oldValue, newValue, startIndex, count);
            return this;
        }

        public override string ToString()
        { 
            return _sb.ToString(); 
        }

        public string ToString(int startIndex, int length)
        {
            return _sb.ToString(startIndex, length);
        }


        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            ISerializable iserializable = _sb;
            iserializable.GetObjectData(info, context);
        }

        Type IStringBuilder.GetType()
        {
           return new Type(GetType());
        }
    }
}
