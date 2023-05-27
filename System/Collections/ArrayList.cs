using GSystem = global::System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System;

namespace DotNetLib.System.Collections
{

    // https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1

    [ComVisible(true)]
    [Description("Implements the IList interface using an array whose size is dynamically increased as required.")]
    [Guid("CE76B55C-7BDD-4471-90B8-704298D3BDF3")]
    [ProgId("DotNetLib.System.Collections.ArrayList")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ArrayList : IArrayList //, ICloneable, GSystem.Collections.IList
    {
        // Constructors

        private GSystem.Collections.ArrayList arrayList;
        public ArrayList()
        {
            arrayList = new GSystem.Collections.ArrayList();
        }

        public ArrayList(int capacity)
        {
            arrayList = new GSystem.Collections.ArrayList(capacity);
        }

        public ArrayList(GSystem.Collections.ICollection c)
        {
            arrayList = new GSystem.Collections.ArrayList(c);
        }

        public ArrayList Create()
        {
            return new ArrayList();
        }

        public ArrayList Create2(int capacity)
        {
            return new ArrayList(capacity);
        }

        public ArrayList Create3(GSystem.Collections.ICollection c)
        {
            return new ArrayList(c);
        }

        public virtual int Capacity
        {
            get { return this.arrayList.Capacity; }
            set { this.arrayList.Capacity = value; }

        }
        public virtual int Count 
        { 
            get { return this.arrayList.Count; }
        }

        public virtual bool IsFixedSize 
        { 
            get { return this.arrayList.IsFixedSize; }
        }


    }
}
