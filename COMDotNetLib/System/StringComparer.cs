// https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparer?view=netframework-4.8.1

using GSystem = global::System;
using GCollections = global::System.Collections;
using DotNetLib.Extensions;
using CultureInfo = DotNetLib.System.Globalization.CultureInfo;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents a string comparison operation that uses specific case and culture-based or ordinal comparison rules.")]
    [Guid("D39835EB-EDAD-4B22-8CB5-E07A11103146")]
    [ProgId("DotNetLib.System.StringComparer")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStringComparer))]
    public class StringComparer  : GCollections.IComparer, GCollections.IEqualityComparer, IWrappedObject, IStringComparer //, IComparer<string>, IEqualityComparer<string>
    {
        private GSystem.StringComparer _stringComparer;

        // Constructors
        public StringComparer(GSystem.StringComparer stringComparer)
        {
            _stringComparer = stringComparer;
        }

        // Properties

        internal GSystem.StringComparer WrappedStringComparer
        {
            get { return _stringComparer; }
            //set { _stringComparer = value; }
        }

        public object WrappedObject => _stringComparer;


        //
        // Summary:
        //     Gets a System.StringComparer object that performs a case-sensitive string comparison
        //     using the word comparison rules of the invariant culture.
        //
        // Returns:
        //     A new System.StringComparer object.
        public static StringComparer InvariantCulture => new StringComparer(GSystem.StringComparer.InvariantCulture);

        //
        // Summary:
        //     Gets a System.StringComparer object that performs a case-insensitive string comparison
        //     using the word comparison rules of the invariant culture.
        //
        // Returns:
        //     A new System.StringComparer object.
        public static StringComparer InvariantCultureIgnoreCase => new StringComparer(GSystem.StringComparer.InvariantCultureIgnoreCase);

        //
        // Summary:
        //     Gets a System.StringComparer object that performs a case-sensitive string comparison
        //     using the word comparison rules of the current culture.
        //
        // Returns:
        //     A new System.StringComparer object.
        public static StringComparer CurrentCulture => new StringComparer(GSystem.StringComparer.CurrentCulture);

        //
        // Summary:
        //     Gets a System.StringComparer object that performs case-insensitive string comparisons
        //     using the word comparison rules of the current culture.
        //
        // Returns:
        //     A new object for string comparison.
        public static StringComparer CurrentCultureIgnoreCase => new StringComparer(GSystem.StringComparer.CurrentCultureIgnoreCase);

        public static StringComparer Ordinal => new StringComparer(GSystem.StringComparer.Ordinal);

        public static StringComparer OrdinalIgnoreCase => new StringComparer(GSystem.StringComparer.OrdinalIgnoreCase);

        public static StringComparer Create(CultureInfo culture, bool ignoreCase)
        {
            return new StringComparer(GSystem.StringComparer.Create(culture.WrappedCultureInfo, ignoreCase));
        }

        public int Compare(object x, object y)
        {
            return _stringComparer.Compare(x, y);
        }

        public int Compare(string x, string y)
        {
            throw new GSystem.NotImplementedException();
        }

        public new bool Equals(object x, object y)
        {
            return _stringComparer.Equals(x, y);
        }

        //
        // Summary:
        //     When overridden in a derived class, indicates whether two strings are equal.
        //
        // Parameters:
        //   x:
        //     A string to compare to y.
        //
        //   y:
        //     A string to compare to x.
        //
        // Returns:
        //     true if x and y refer to the same object, or x and y are equal, or x and y are
        //     null; otherwise, false.
        public bool Equals(string x, string y)
        {
            return _stringComparer.Equals(x, y);

        }

        public int GetHashCode(object obj)
        {
            return _stringComparer.GetHashCode(obj);
        }

        public int GetHashCode(string obj)
        {
            return _stringComparer.GetHashCode(obj);
        }

        new public Type GetType()
        {
            return new Type(((object)this).GetType());
        }
    }
}
