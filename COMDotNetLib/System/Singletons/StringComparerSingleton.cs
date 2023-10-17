// https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparer?view=netframework-4.8.1

using DotNetLib.System.Globalization;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Description("Represents a string comparison operation that uses specific case and culture-based or ordinal comparison rules.")]
    [Guid("477806D7-3A2D-4A1A-9378-B00AA634F7BD")]
    [ProgId("DotNetLib.System.StringComparerSingleton")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IStringComparerSingleton))]
    public class StringComparerSingleton : IStringComparerSingleton
    {
        public StringComparerSingleton() { }

        // Static properites
        public StringComparer InvariantCulture => StringComparer.InvariantCulture;

        public StringComparer InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase;

        public StringComparer CurrentCulture => StringComparer.CurrentCulture;

        public StringComparer CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase;

        public StringComparer Ordinal => StringComparer.Ordinal;

        public StringComparer OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase;


        // Methods

        public StringComparer Create(CultureInfo culture, bool ignoreCase)
        {
            return StringComparer.Create(culture, ignoreCase);

        }

    }
}
