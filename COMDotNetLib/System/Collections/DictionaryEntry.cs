// https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using GCollections = global::System.Collections;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Defines a dictionary key/value pair that can be set or retrieved.")]
    [Guid("B0E606C8-0FCA-4395-B90A-142D7C622ED9")]
    [ProgId("DotNetLib.System.Collections.DictionaryEntry")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDictionaryEntry))]
    public class DictionaryEntry : IDictionaryEntry
    {
        private GCollections.DictionaryEntry _dictionaryEntry;

        public DictionaryEntry(GCollections.DictionaryEntry dictionaryEntry)
        {
            _dictionaryEntry = dictionaryEntry;
        }

        public DictionaryEntry(object key, object value)
        {
            _dictionaryEntry = new GCollections.DictionaryEntry(key, value);
        }
        public object Key
        {
            get => _dictionaryEntry.Key;
            set => _dictionaryEntry.Key = value;
        }

        public object Value 
        {
            get => _dictionaryEntry.Value;
            set => _dictionaryEntry.Value = value;
        }

        // Added to obtain the System.Collection.DictionaryEntry which is a struc, value type
        public void GetDictionaryEntry([In][Out] ref GCollections.DictionaryEntry dictEntry)
        {
            dictEntry = _dictionaryEntry;
        }
    }
}
