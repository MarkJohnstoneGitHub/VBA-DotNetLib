// https://coordinatesharp.com/Help/html/T_CoordinateSharp_Coordinate.htm

using GCSharp = global::CoordinateSharp; 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DotNetLib.CoordinateSharp
{
    public class Coordinate
    {
        private GCSharp.Coordinate _coordinate;

        public Coordinate() 
        { 
            _coordinate = new GCSharp.Coordinate();
        }

        public Coordinate(GCSharp.Coordinate coordinate) 
        { 
            _coordinate = coordinate;
        }

        public Coordinate(double lat, double longi)
        {
            _coordinate = new GCSharp.Coordinate(lat, longi);
        }

        public Coordinate(double lat, double longi, DateTime date)
        {
            _coordinate = new GCSharp.Coordinate(lat, longi, date);
        }





    }
}
