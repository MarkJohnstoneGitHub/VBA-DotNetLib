// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("0E55839B-EF94-42F1-9AF2-B1C0F97EF372")]
    [ProgId("DotNetLib.System.Globalization.CompareInfoSingleton")]
    [Description("")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICompareInfoSingleton))]
    public class CompareInfoSingleton : ICompareInfoSingleton
    {
        public CompareInfoSingleton() {}

        public CompareInfo GetCompareInfo(string name)
        {
            return CompareInfo.GetCompareInfo(name);
        }
        public  CompareInfo GetCompareInfo2(int culture)
        {
            return CompareInfo.GetCompareInfo(culture);
        }

        public bool IsSortable(string text)
        {
            return CompareInfo.IsSortable(text);
        }
    }
}
