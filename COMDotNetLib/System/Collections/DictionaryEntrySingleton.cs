// https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Defines a dictionary key/value pair that can be set or retrieved.")]
    [Guid("70DD851D-95F9-42BC-9E5B-1979070FF6C7")]
    [ProgId("DotNetLib.System.Collections.DictionaryEntrySingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDictionaryEntrySingleton))]
    public class DictionaryEntrySingleton : IDictionaryEntrySingleton
    {
        public DictionaryEntry Create(object key, object value)
        {
            return new DictionaryEntry(key, value);
        }

        public DictionaryEntry Create(GCollections.DictionaryEntry dictionaryEntry)
        {
            return new DictionaryEntry(dictionaryEntry);
        }


    }
}
