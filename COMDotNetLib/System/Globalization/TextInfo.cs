// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1

using GGlobalization = global::System.Globalization;

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [Serializable]

    [ComVisible(true)]
    [Guid("8338BC0F-C6EA-49A1-9644-6A0D57FDD1CA")]
    [ProgId("DotNetLib.System.TextInfo")]
    [Description("Defines text properties and behaviors, such as casing, that are specific to a writing system.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ITextInfo))]

    public class TextInfo : ICloneable, ITextInfo
    {

        private GGlobalization.TextInfo _textInfo;

        internal TextInfo(GGlobalization.TextInfo textInfo)
        {
            _textInfo = textInfo;
        }

        internal GGlobalization.TextInfo TextInfoObject
        {
            get => _textInfo;
            set => _textInfo = value;
        }

        public int ANSICodePage => _textInfo.ANSICodePage;

        public string CultureName => _textInfo.CultureName;

        public int EBCDICCodePage => _textInfo.EBCDICCodePage;

        public bool IsReadOnly => _textInfo.IsReadOnly;

        public bool IsRightToLeft => _textInfo.IsRightToLeft;

        public int LCID => _textInfo.LCID;

        public string ListSeparator
        {
            get => _textInfo.ListSeparator;
            set => _textInfo.ListSeparator = value;
        }

        public int MacCodePage => _textInfo.MacCodePage;

        public int OEMCodePage => _textInfo.OEMCodePage;

        public object Clone()
        {
            return new TextInfo((GGlobalization.TextInfo)_textInfo.Clone());
         
        }

        public override bool Equals(object obj)
        {
            TextInfo textInfo = obj as TextInfo;
            if (textInfo != null)
            {
                return CultureName.Equals(textInfo.CultureName);
            }
            return false;
        }

        public override int GetHashCode()
        {
            return _textInfo.GetHashCode();
        }

        public static TextInfo ReadOnly(TextInfo textInfo)
        {
            return new TextInfo(GGlobalization.TextInfo.ReadOnly(textInfo.TextInfoObject));
        }
        public string ToLower(string str)
        {
            return _textInfo.ToLower(str);
        }

        public new string ToString()
        {
            return _textInfo.ToString();
        }

        public string ToTitleCase(string str)
        {
            return _textInfo.ToTitleCase(str);
        }

        public string ToUpper(string str)
        {
            return _textInfo.ToUpper(str);    
        }

    }
}
