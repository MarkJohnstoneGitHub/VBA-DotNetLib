// https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.System.Globalization
{
    [ComVisible(true)]
    [Guid("41D4BF7E-71D2-42C1-9F8B-A9170A1D308B")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICompareInfoSingleton
    {
        [Description("Initializes a new CompareInfo object that is associated with the culture with the specified identifier.")]
        CompareInfo GetCompareInfo(string name);

        [Description("Initializes a new CompareInfo object that is associated with the culture with the specified name.")]
        CompareInfo GetCompareInfo2(int culture);

        [Description("Indicates whether a specified Unicode string is sortable.")]
        bool IsSortable(string text);


    }
}
