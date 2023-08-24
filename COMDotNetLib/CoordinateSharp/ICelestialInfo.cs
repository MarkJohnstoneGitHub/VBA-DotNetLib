// https://www.nuget.org/packages/CoordinateSharp/#readme-body-tab
// https://coordinatesharp.com/Help/html/T_CoordinateSharp_Celestial.htm

using DateTime = DotNetLib.System.DateTime;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;


namespace DotNetLib.CoordinateSharp
{
    [ComVisible(true)]
    [Guid("AF12D88C-4CCD-4434-A7B3-26FB5BD97206")]
    [Description("")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICelestialInfo
    {
        // Methods

        [Description("Calcualtes sunrise time.")]
        DateTime SunRise(double lat, double longi, DateTime date);

        [Description("Calculates sunset time.")]
        DateTime SunSet(double lat, double longi, DateTime date);
    }
}
