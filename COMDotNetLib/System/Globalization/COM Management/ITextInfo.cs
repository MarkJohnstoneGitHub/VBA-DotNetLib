// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("0E20DC25-693E-4F24-87E0-0ABBE5998D9F")]
    [Description("Defines text properties and behaviors, such as casing, that are specific to a writing system.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ITextInfo  //,IDeserializationCallback
    {
        // Properties
        int ANSICodePage 
        {
            [Description("Gets the American National Standards Institute (ANSI) code page used by the writing system represented by the current TextInfo.")]
            get;
        }

        string CultureName
        {
            [Description("")]
            get;
        }

        int EBCDICCodePage 
        {
            [Description("Gets the name of the culture associated with the current TextInfo object.")]
            get;
        }

        bool IsReadOnly 
        {
            [Description("Gets a value indicating whether the current TextInfo object is read-only.")]
            get;
        }

        bool IsRightToLeft 
        {
            [Description("Gets a value indicating whether the current TextInfo object represents a writing system where text flows from right to left.")]
            get;
        }

        int LCID 
        {
            [Description("Gets the culture identifier for the culture associated with the current TextInfo object.")]
            get;
        }

        string ListSeparator
        {
            [Description("Gets or sets the string that separates items in a list.")]
            get;
            [Description("Gets or sets the string that separates items in a list.")]
            set;
        }

        int MacCodePage 
        {
            [Description("Gets the Macintosh code page used by the writing system represented by the current TextInfo.")]
            get;
        }

        int OEMCodePage 
        {
            [Description("Gets the original equipment manufacturer (OEM) code page used by the writing system represented by the current TextInfo.")]
            get;
        }

        //Methods
        [Description("Creates a new object that is a copy of the current TextInfo object.")]
        object Clone();

        [Description("Determines whether the specified object represents the same writing system as the current TextInfo object.")]
        bool Equals(object obj);

        [Description("Serves as a hash function for the current TextInfo, suitable for hashing algorithms and data structures, such as a hash table.")]
        int GetHashCode();

        [Description("Converts the specified string to lowercase")]
        string ToLower(string str);

        [Description("Returns a string that represents the current TextInfo.")]
        string ToString();

        [Description("")]
        string ToTitleCase(string str);

        [Description("Converts the specified string to uppercase.")]
        string ToUpper(string str);
        //new void OnDeserialization(object sender);

    }
}
