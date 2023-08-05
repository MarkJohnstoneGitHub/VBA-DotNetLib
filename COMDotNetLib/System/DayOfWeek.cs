// https://learn.microsoft.com/en-us/dotnet/api/system.dayofweek?view=netframework-4.8.1

using System.Runtime.InteropServices;

namespace DotNetLib.System
{
    [ComVisible(true)]
    [Guid("6E5F9410-9B91-4CA9-B6D0-4C77FEC49B00")]
    public enum DayOfWeek
    {
        Sunday = 0,
        Monday = 1,
        Tuesday = 2,
        Wednesday = 3,
        Thursday = 4,
        Friday = 5,
        Saturday = 6,
    }

    // Defining IDL Enumerations
    // https://www.akella.org/homepages/mani/documents/secdocs/ASP.net/ComAndDotNetInteroperability.pdf
}
