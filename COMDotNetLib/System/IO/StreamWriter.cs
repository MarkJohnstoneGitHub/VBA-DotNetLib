// https://learn.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8.1

using GTasks = global::System.Threading.Tasks;
using GSystem = global::System;
using GIO = global::System.IO;
using GText = global::System.Text;
using GRemoting = global::System.Runtime.Remoting;
using System;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using DotNetLib.Extensions;
using System.ComponentModel;

namespace DotNetLib.System.IO
{
    [ComVisible(true)]
    [Description("Implements a TextWriter for writing characters to a stream in a particular encoding.")]
    [Guid("877CFF89-EDE5-40D4-9674-588EEF3A6203")]
    [ProgId("DotNetLib.System.IO.StreamWriter")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStreamWriter))]
    public class StreamWriter : IStreamWriter, IWrappedObject, IDisposable
    {
        private GIO.StreamWriter _streamWriter;

        public static readonly StreamWriter Null = new StreamWriter(GIO.StreamWriter.Null);

        internal StreamWriter(GIO.StreamWriter streamWriter)
        {
            _streamWriter = streamWriter;
        }

        public StreamWriter(string path)
        {
            _streamWriter = new GIO.StreamWriter(path);
        }

        public StreamWriter(string path, bool append)
        {
            _streamWriter = new GIO.StreamWriter(path, append);
        }

        public StreamWriter(string path, bool append, GText.Encoding encoding)
        {
            _streamWriter = new GIO.StreamWriter(path, append, encoding);
        }

        public StreamWriter(string path, bool append, GText.Encoding encoding, int bufferSize)
        {
            _streamWriter = new GIO.StreamWriter(path, append, encoding, bufferSize);
        }

        public StreamWriter(GIO.Stream stream)
        {
            _streamWriter = new GIO.StreamWriter(stream);
        }

        public StreamWriter(GIO.Stream stream, GText.Encoding encoding)
        {
            _streamWriter = new GIO.StreamWriter(stream,encoding);
        }

        public StreamWriter(GIO.Stream stream, GText.Encoding encoding, int bufferSize)
        {
            _streamWriter = new GIO.StreamWriter(stream, encoding, bufferSize);
        }

        public StreamWriter(GIO.Stream stream, GText.Encoding encoding, int bufferSize, bool leaveOpen)
        {
            _streamWriter = new GIO.StreamWriter(stream, encoding, bufferSize, leaveOpen);
        }

        // Properties

        public virtual bool AutoFlush 
        {
            get => _streamWriter.AutoFlush;
            set => _streamWriter.AutoFlush = value;
        }

        public virtual GIO.Stream BaseStream => _streamWriter.BaseStream;

        public GText.Encoding Encoding => _streamWriter.Encoding;

        public  IFormatProvider FormatProvider => _streamWriter.FormatProvider;

        public string NewLine
        {
            get => _streamWriter.NewLine;
            set => _streamWriter.NewLine = value;
        }

        public object WrappedObject => _streamWriter;

        // Methods

        public void Close()
        { 
            _streamWriter.Close(); 
        }

        public virtual GRemoting.ObjRef CreateObjRef(Type requestedType)
        {
            return _streamWriter.CreateObjRef(requestedType.WrappedType);
        }

        public void Dispose()
        { 
            _streamWriter.Dispose(); 
        }

        public new virtual bool Equals(object obj)
        {
            return _streamWriter.Equals(obj.Unwrap());
        }
        public void Flush()
        {
            _streamWriter.Flush();
        }

        public GTasks.Task FlushAsync()
        {
            return _streamWriter.FlushAsync();
        }

        public new virtual int GetHashCode()
        { 
            return _streamWriter.GetHashCode(); 
        }

        public object GetLifetimeService()
        {
            return _streamWriter.GetLifetimeService();
        }

        public new Type GetType()
        {
            return new Type(((GSystem.Object)this).GetType());
        }

        public virtual object InitializeLifetimeService()
        {
            return _streamWriter.InitializeLifetimeService();
        }

        public new virtual string ToString()
        { 
            return _streamWriter.ToString(); 
        }

        public void Write(string value)
        { 
            _streamWriter.Write(value); 
        }

        public virtual void Write2(bool value)
        {
            _streamWriter.Write(value);
        }

        public virtual void Write3(int value)
        { 
            _streamWriter.Write(value);
        }

        public virtual void Write4(long value)
        { 
            _streamWriter.Write(value);
        }

        public void Write5(float value)
        {
            _streamWriter.Write(value);
        }
        public virtual void Write6(double value)
        {
            _streamWriter.Write(value);
        }

        public virtual void Write7(object value)
        {
            _streamWriter.Write(value.Unwrap());
        }

        public virtual void Write8(string format, object arg0)
        {
            _streamWriter.Write(format, arg0.Unwrap());
        }

        public virtual void Write9(string format, object arg0, object arg1)
        {
            _streamWriter.Write(format, arg0.Unwrap(), arg1.Unwrap());
        }

        public virtual void Write10(string format, object arg0, object arg1, object arg2)
        {
            _streamWriter.Write(format, arg0.Unwrap(), arg1.Unwrap(), arg2.Unwrap());
        }

        public virtual void Write11(string format, [In] ref object[] arg)
        {
            _streamWriter.Write(format, arg.Unwrap());
        }


        public virtual void WriteLine()
        { 
            _streamWriter.WriteLine(); 
        }

        public virtual void WriteLine2(string value)
        { 
            _streamWriter.WriteLine(value); 
        }

        public virtual void WriteLine3(bool value)
        { 
            _streamWriter.WriteLine(value); 
        }

        public virtual void WriteLine4(int value)
        { 
            _streamWriter.WriteLine(value);
        }

        public virtual void WriteLine5(long value)
        { 
            _streamWriter.WriteLine(value);
        }

        public virtual void WriteLine6(float value)
        {
            _streamWriter.WriteLine(value);
        }
        public virtual void WriteLine7(double value)
        {
            _streamWriter.WriteLine(value);
        }

        public virtual void WriteLine8(object value)
        { 
            _streamWriter.WriteLine(value.Unwrap());
        }

        public GTasks.Task WriteLineAsync()
        { 
            return _streamWriter.WriteLineAsync(); 
        }

        public GTasks.Task WriteLineAsync(string value)
        { 
            return _streamWriter.WriteLineAsync(value); 
        }


        //public virtual void WriteLine2(char[] buffer)
        //{ 
        //    _streamWriter.WriteLine2(buffer);
        //}

        //public virtual void WriteLine2(StringBuilder sb)
        //{ 
        //    _streamWriter.WriteLine2(sb.WrappedStringBuilder); 
        //}


    }
}
