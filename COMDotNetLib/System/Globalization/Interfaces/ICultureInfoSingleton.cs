using GGlobalization = global::System.Globalization;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("33F77CBC-6975-4953-82D6-DF4DD1C5EA3E")]
    [Description("Provides information about a specific culture (called a locale for unmanaged code development). The information includes the names for the culture, the writing system, the calendar used, the sort order of strings, and formatting for dates and numbers.")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICultureInfoSingleton
    {
        // Constructors

        [Description("Initializes a new instance of the CultureInfo class based on the culture specified by the culture identifier.")]
        CultureInfo Create(int culture);

        [Description("Initializes a new instance of the CultureInfo class based on the culture specified by name.")]
        CultureInfo Create2(string name);

        [Description("Initializes a new instance of the CultureInfo class based on the culture specified by the culture identifier and on a value that specifies whether to use the user-selected culture settings from Windows.")]
        CultureInfo Create3(int culture, bool useUserOverride);

        [Description("Initializes a new instance of the CultureInfo class based on the culture specified by name and on a value that specifies whether to use the user-selected culture settings from Windows.")]
        CultureInfo Create4(string name, bool useUserOverride);

        // Properties

        CultureInfo CurrentCulture
        {
            [Description("Gets or sets the CultureInfo object that represents the culture used by the current thread and task-based asynchronous operations.")]
            get;
            [Description("Gets or sets the CultureInfo object that represents the culture used by the current thread and task-based asynchronous operations.")]
            set;
        }

        CultureInfo CurrentUICulture
        {
            [Description("Gets or sets the CultureInfo object that represents the current user interface culture used by the Resource Manager to look up culture-specific resources at run time.")]
            get;
            [Description("Gets or sets the CultureInfo object that represents the current user interface culture used by the Resource Manager to look up culture-specific resources at run time.")]
            set;
        }

        CultureInfo DefaultThreadCurrentCulture
        {
            [Description("Gets or sets the default culture for threads in the current application domain.")]
            get;
            [Description("Gets or sets the default culture for threads in the current application domain.")]
            set;
        }

        CultureInfo DefaultThreadCurrentUICulture
        {
            [Description("Gets or sets the default UI culture for threads in the current application domain.")]
            get;
            [Description("Gets or sets the default UI culture for threads in the current application domain.")]
            set;
        }

        CultureInfo InstalledUICulture
        {
            [Description("Gets the CultureInfo that represents the culture installed with the operating system.")]
            get;
        }

        CultureInfo InvariantCulture
        {
            [Description("Gets the CultureInfo object that is culture-independent(invariant).")]
            get;
        }

        // Methods

        [Description("Creates a CultureInfo that represents the specific culture that is associated with the specified name.")]
        CultureInfo CreateSpecificCulture(string name);

        // GetCultureInfo Overloads
        [Description("Retrieves a cached, read-only instance of a culture by using the specified culture identifier.")]
        CultureInfo GetCultureInfo(int culture);

        [Description("Retrieves a cached, read-only instance of a culture using the specified culture name.")]
        CultureInfo GetCultureInfo2(string name);

        [Description("Retrieves a cached, read-only instance of a culture. Parameters specify a culture that is initialized with the TextInfo and CompareInfo objects specified by another culture.")]
        CultureInfo GetCultureInfo3(string name, string altName);

        [Description("Deprecated. Retrieves a read-only CultureInfo object having linguistic characteristics that are identified by the specified RFC 4646 language tag.")]
        CultureInfo GetCultureInfoByIetfLanguageTag(string name);

        [Description("Gets the list of supported cultures filtered by the specified CultureTypes parameter.")]
        CultureInfo[] GetCultures(GGlobalization.CultureTypes types);

        [Description("Returns a read-only wrapper around the specified CultureInfo object.")] 
        CultureInfo ReadOnly(CultureInfo ci);
    }
}
