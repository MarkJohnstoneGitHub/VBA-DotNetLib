﻿// https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Guid("9135A43E-C615-4B2D-A820-038026E2F26A")]
    [Description("Defines a dictionary key/value pair that can be set or retrieved.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDictionaryEntrySingleton
    {
        //[Description("Initializes an instance of the DictionaryEntry type from an object containing a mscorlib.DictionaryEntry")]
        //DictionaryEntry Create2(object dictionaryEntry);

        //[Description("Initializes an instance of the DictionaryEntry type from a mscorlib.DictionaryEntry containing the specified key and value.")]
        //DictionaryEntry Create2([In] ref GCollections.DictionaryEntry dictionaryEntry);

        //[Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")]
        //DictionaryEntry Create2(object key, object value);

        //[Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")]
        //GCollections.DictionaryEntry Create2(object key, object value);

        [Description("Initializes an instance of the mscorlib.DictionaryEntry type with the specified key and value.")]
        void Create([In][Out] ref GCollections.DictionaryEntry dictionaryEntry, object key, object value);

        [Description("Gets the key in the key/value pair.")]
        object GetKey([In] ref GCollections.DictionaryEntry dictionaryEntry);

        [Description("Gets the value in the key/value pair.")]
        object GetValue([In] ref GCollections.DictionaryEntry dictionaryEntry);

        [Description("Sets the key in the key/value pair.")]
        void SetKey([In] ref GCollections.DictionaryEntry dictionaryEntry, object key);

        void SetValue([In] ref GCollections.DictionaryEntry dictionaryEntry, object value);

    }
}
