// https://learn.microsoft.com/en-us/dotnet/api/system.predicate-1?view=netframework-4.8.1

using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("E4CB5D3C-540B-40BE-8B8F-7C1129E26AC0")]
    [ProgId("DotNetLib.System.PredicateSingleton")]
    [Description("Represents the method that defines a set of criteria and determines whether the specified object meets those criteria.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IPredicateSingleton))]
    public class PredicateSingleton : IPredicateSingleton
    {
        public PredicateSingleton() { }

        // Factory Methods
        public Predicate Create(IPredicate predicate)
        {
            return new Predicate(predicate);
        }

    }
}
