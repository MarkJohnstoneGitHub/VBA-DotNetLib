// https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

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
        [Description("Initializes an instance of the DictionaryEntry type with the specified key and value.")]
        DictionaryEntry Create(object key, object value);

        [Description("Initializes an instance of the DictionaryEntry type from a mscorlib.DictionaryEntry containing the specified key and value.")]
        DictionaryEntry Create(GCollections.DictionaryEntry dictionaryEntry);
    }
}
