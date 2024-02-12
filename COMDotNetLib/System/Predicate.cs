// https://learn.microsoft.com/en-us/dotnet/api/system.predicate-1?view=netframework-4.8.1
// https://stackoverflow.com/questions/556425/what-is-a-predicate-delegate-and-where-should-it-be-used
// https://stackoverflow.com/a/556442/10759363
// https://web.archive.org/web/20140625132124/http://blog.coverity.com/2014/06/18/delegates-structural-identity/#.U6rM93bP1qY

using GSystem = global::System;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("6AE61D70-C6A6-40E9-B645-7DEFE19A2D15")]
    [ProgId("DotNetLib.System.Predicate")]
    [Description("Represents the method that defines a set of criteria and determines whether the specified object meets those criteria.")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IPredicate))]
    public class Predicate : IPredicate
    {
        internal IPredicate _predicate;

        // https://stackoverflow.com/questions/731249/how-to-convert-funct-bool-to-predicatet
        //internal Func<object, bool> FuncPredicate;
        //internal GSystem.CallBack<object> _myPredicate;

        public Predicate(IPredicate predicate)
        {
            if (predicate == null)
            {
                throw new ArgumentNullException("predicate");
            }
            _predicate = predicate;
        }

        public bool CallBack(object obj)
        {
            //if (_predicate == null)
            //{
            //    throw new ArgumentNullException("predicate");
            //}
            return _predicate.CallBack(obj);
        }

        //Call back to non-generic VBA predicate?
        internal bool PredicateCallback<T>(T obj)
        {
            //if (_predicate == null)
            //{
            //    throw new ArgumentNullException("predicate");
            //}
            return _predicate.CallBack(obj);
        }

    }
}
