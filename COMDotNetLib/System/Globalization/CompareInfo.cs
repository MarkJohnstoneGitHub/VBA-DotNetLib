// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;
using System.Globalization;
using DotNetLib.Extensions;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("2A987B84-CBFA-4A57-A946-0CA6B5B69FF3")]
    [ProgId("DotNetLib.System.Globalization.CompareInfo")]
    [Description("Implements a set of methods for culture-sensitive string comparisons.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICompareInfo))]

    public class CompareInfo : ICompareInfo, IWrappedObject
    {
        private GGlobalization.CompareInfo _compareInfo;

        public CompareInfo(GGlobalization.CompareInfo compareInfo)
        {
            WrappedCompareInfo = compareInfo;
        }

        // Properties
        internal GGlobalization.CompareInfo WrappedCompareInfo
        {
            get { return _compareInfo; }
            set { _compareInfo = value; }  
        }

        public object WrappedObject => _compareInfo;

        public int LCID => _compareInfo.LCID; 
        
        public string Name => _compareInfo.Name;

        public GGlobalization.SortVersion Version => _compareInfo.Version;

        // Methods
        public int Compare(string string1, string string2, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.Compare(string1, string2, options);
        }

        public int Compare2(string string1, int offset1, string string2, int offset2, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.Compare(string1, offset1, string2, offset2, options);
        }

        public int Compare3(string string1, int offset1, int length1, string string2, int offset2, int length2, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.Compare(string1, offset1,length1,string2,offset2,length2,options);
        }

        public override bool Equals(object value)
        {
            return value is CompareInfo ci && _compareInfo.Equals(ci.WrappedCompareInfo);
        }

        public override int GetHashCode()
        {
            return _compareInfo.GetHashCode();
        }

        public static CompareInfo GetCompareInfo(int culture)
        {
            return new CompareInfo(GGlobalization.CompareInfo.GetCompareInfo(culture));
        }

        public static CompareInfo GetCompareInfo(string name)
        {
            return new CompareInfo(GGlobalization.CompareInfo.GetCompareInfo(name));
        }

        public GGlobalization.SortKey GetSortKey(string source, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.GetSortKey(source,options);
        }

        //public int IndexOf(string source, string value, CompareOptions options = CompareOptions.None)
        //{
        //    return _compareInfo.IndexOf(source,value, options);
        //}

        public int IndexOf(string source, string value, int startIndex = 0, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.IndexOf(source, value, startIndex, options);
        }

        public int IndexOf2(string source, string value, int startIndex, int count, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.IndexOf(source, value, startIndex, count, options);
        }

        public bool IsPrefix(string source, string prefix, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.IsPrefix(source,prefix,options);
        }

        public static bool IsSortable(string text)
        {
            return GGlobalization.CompareInfo.IsSortable(text);
        }

        public bool IsSuffix(string source, string suffix, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.IsSuffix(source,suffix,options);
        }

        public int LastIndexOf(string source, string value, int startIndex = 0, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.LastIndexOf(source, value, startIndex, options);
        }

        public int LastIndexOf2(string source, string value, int startIndex, int count, CompareOptions options = CompareOptions.None)
        {
            return _compareInfo.LastIndexOf(source, value, startIndex, count, options);
        }

        public override string ToString()
        { 
            return _compareInfo.ToString(); 
        }

    }
}
