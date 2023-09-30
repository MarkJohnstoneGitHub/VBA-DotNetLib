// https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capture?view=netframework-4.8.1

using GRegularExpressions = global::System.Text.RegularExpressions;
using DotNetLib.Extensions;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Reflection;

namespace DotNetLib.System.Text.RegularExpressions
{
    [ComVisible(true)]
    [Guid("2D051F1C-2269-46BC-8433-D196E4075E66")]
    [ProgId("DotNetLib.System.Text.RegularExpressions.Capture")]
    [Description("Represents the results from a single successful subexpression capture.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICapture))]
    public class Capture : ICapture , IWrappedObject
    {
        internal GRegularExpressions.Capture _capture;
        //internal Group group;

        internal Capture(GRegularExpressions.Capture capture) 
        { 
            _capture = capture;
        }

        // Properties

        public object WrappedObject => _capture;

        internal GRegularExpressions.Capture WrappedCapture => _capture;

        public int Index => _capture.Index;

        public int Length => _capture.Length;

        public string Value => _capture.Value;

        // Methods

        //Todo check implementation
        public new virtual bool Equals(object obj)
        {
            return Equals(obj.Unwrap() as GRegularExpressions.Capture);
        }

        public new virtual int GetHashCode()
        { 
            return _capture.GetHashCode(); 
        }

        public override string ToString()
        {  
            return _capture.ToString();
        }

    }
}
