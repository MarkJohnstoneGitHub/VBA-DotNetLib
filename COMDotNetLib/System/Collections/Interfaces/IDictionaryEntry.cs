// https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1

//Note: No longer in use as using DictionaryEntrySingleton passing by reference

using GCollections = global::System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Collections
{
    [ComVisible(false)]
    [Guid("0CD703F2-7AA7-464C-AC21-C2C97EB47171")]
    [Description("Defines a dictionary key/value pair that can be set or retrieved.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDictionaryEntry
    {
        object Key 
        {
            [Description("Gets or sets the key in the key/value pair.")]
            get;
            [Description("Gets or sets the key in the key/value pair.")]
            set;
        }

        object Value 
        {
            [Description("Gets or sets the value in the key/value pair.")]
            get;
            [Description("Gets or sets the value in the key/value pair.")]
            set;
        }

        // Added to obtain the System.Collection.DictionaryEntry which is a struc, value type
        [Description("Gets a mscorlib.DictionaryEntry value type for a key/value pair.")]
        void GetDictionaryEntry([In][Out] ref GCollections.DictionaryEntry dictEntry);

    }
}
