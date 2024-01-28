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

        public StringBuilder Append2(string value, int startIndex, int count)
        {
            _sb.Append(value, startIndex, count);
            return this;
        }

        public StringBuilder Append(StringBuilder value)
        {
            _sb.Append(value.WrappedStringBuilder);
            return this;
        }

        public StringBuilder Append3(bool value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append4(byte value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append5(short value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append6(int value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append7(long value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append8(float value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append9(double value)
        {
            _sb.Append(value);
            return this;
        }

        //public StringBuilder Append2(decimal value)
        //{
        //    _sb.Append2(value);
        //    return this;
        //}

        //Todo Check implementation  is Unwrap() required?
        public StringBuilder Append10(object value)
        {
            _sb.Append(value);
            return this;
        }

        public StringBuilder Append11(String value)
        {
            _sb.Append(value.WrappedString);
            return this;
        }

        public StringBuilder Append12(String value, int startIndex, int count)
        {
            _sb.Append(value.WrappedString, startIndex, count);
            return this;
        }

        public StringBuilder AppendFormat(string format, object arg0)
        {
            _sb.AppendFormat(format, arg0);
            return this;
        }

        public StringBuilder AppendFormat2(string format, object arg0, object arg1)
        {
            _sb.AppendFormat(format, arg0, arg1);
            return this;
        }

        public StringBuilder AppendFormat3(string format, object arg0, object arg1, object arg2)
        {
            _sb.AppendFormat(format, arg0, arg1, arg2);
            return this;
        }

        public StringBuilder AppendFormat4(string format, [In] ref object[] args)
        {
            _sb.AppendFormat(format, args);
            return this;
        }

        public StringBuilder AppendFormat5(IFormatProvider provider, string format, object arg0)
        {
            _sb.AppendFormat(provider.Unwrap(), format, arg0);
            return this;
        }

        public StringBuilder AppendFormat6(IFormatProvider provider, string format, object arg0, object arg1)
        {
            _sb.AppendFormat(provider.Unwrap(), format, arg0, arg1);
            return this;
        }

        public StringBuilder AppendFormat7(IFormatProvider provider, string format, object arg0, object arg1, object arg2)
        {
            _sb.AppendFormat(provider.Unwrap(), format, arg0, arg1, arg2);
            return this;
        }

        public StringBuilder AppendFormat8(IFormatProvider provider, string pFormat, [In] ref object[] args)
        {
            _sb.AppendFormat(provider.Unwrap(), pFormat, args);
            return this;
        }


        public StringBuilder AppendLine()
        {
            _sb.AppendLine();
            return this;
        }

        public StringBuilder AppendLine2(string value)
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

        public StringBuilder Insert2(int index, string value, int count)
        {
            _sb.Insert(index, value, count);
            return this;
        }

        public StringBuilder Insert3(int index, bool value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert4(int index, byte value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert5(int index, short value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert6(int index, int value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert7(int index, long value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert8(int index, float value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert9(int index, double value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert(int index, decimal value)
        {
            _sb.Insert(index, value);
            return this;
        }

        public StringBuilder Insert10(int index, object value)
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

        public StringBuilder Replace2(string oldValue, string newValue, int startIndex, int count)
        {
            _sb.Replace(oldValue, newValue, startIndex, count);
            return this;
        }

        public override string ToString()
        { 
            return _sb.ToString(); 
        }

        public string ToString2(int startIndex, int length)
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
