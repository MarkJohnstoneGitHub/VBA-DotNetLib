// https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1

using GCollections = global::System.Collections;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.Collections
{
    [ComVisible(true)]
    [Description("Represents a collection of key/value pairs that are organized based on the hash code of the key.")]
    [Guid("B636CE48-6D0B-4C3F-A09B-EAE25245C981")]
    [ProgId("DotNetLib.System.Collections.HashtableSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IHashtableSingleton))]
    public class HashtableSingleton : IHashtableSingleton
    {
        //public Hashtable Create2(int capacity = 0, float loadFactor = 1f)
        //{
        //    return new Hashtable(capacity, loadFactor);
        //}

        public Hashtable Create(int capacity = 0, float loadFactor = 1f, GCollections.IEqualityComparer equalityComparer = null)
        {
            return new Hashtable(capacity, loadFactor, equalityComparer);
        }

        public Hashtable Create2(GCollections.IDictionary d, float loadFactor = 1f, GCollections.IEqualityComparer equalityComparer = null)
        {
            return new Hashtable(d, loadFactor, equalityComparer);
        }

    }
}
