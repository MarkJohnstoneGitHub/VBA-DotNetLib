// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("448E87E5-8496-4FCA-8DDB-0D53AC21577E")]
    [ProgId("DotNetLib.System.Globalization.TextInfoSingleton")]
    [Description("TextInfo factory methods and static members.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITextInfoSingleton))]

    public class TextInfoSingleton : ITextInfoSingleton
    {
        public TextInfoSingleton() { }

        public TextInfo ReadOnly(TextInfo textInfo)
        {
            return TextInfo.ReadOnly(textInfo);
        }

    }
}
