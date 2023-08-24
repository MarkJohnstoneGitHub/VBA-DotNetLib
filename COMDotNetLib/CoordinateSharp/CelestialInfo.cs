// // https://coordinatesharp.com/Help/html/T_CoordinateSharp_Celestial.htm

using CoordinateSharp; // https://www.nuget.org/packages/CoordinateSharp/#readme-body-tab
using System;
using DateTime = DotNetLib.System.DateTime;
using System.ComponentModel;
using System.Runtime.InteropServices;
using GCSharp = global::CoordinateSharp;


namespace DotNetLib.CoordinateSharp
{

    [ComVisible(true)]
    [Guid("9908BDC6-10C5-47E6-89A6-3D711848C81D")]
    [ProgId("CoordinateSharp.")]
    [Description("")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICelestialInfo))]
    public class CelestialInfo : ICelestialInfo
    {
        public CelestialInfo()
        {
        }

        public DateTime SunRise(double lat, double longi, DateTime date)
        {
            //GCSharp.Coordinate coordinate = new GCSharp.Coordinate();

            GCSharp.Coordinate c = new GCSharp.Coordinate(lat, longi, date.DateTimeObject);

            return new DateTime(c.CelestialInfo.SunRise.Value);
        }

        public DateTime SunSet(double lat, double longi, DateTime date)
        {
            GCSharp.Coordinate c = new GCSharp.Coordinate(lat, longi, date.DateTimeObject);

            return new DateTime(c.CelestialInfo.SunSet.Value);
        }
    }
}

