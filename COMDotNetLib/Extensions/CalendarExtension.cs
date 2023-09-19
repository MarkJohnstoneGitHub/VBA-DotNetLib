using DotNetLib.System.Globalization;
using GGlobalization = global::System.Globalization;
using System.Runtime.InteropServices;
//using GregorianCalendar = DotNetLib.System.Globalization.GregorianCalendar;
//using PersianCalendar = DotNetLib.System.Globalization.PersianCalendar;
//using UmAlQuraCalendar = DotNetLib.System.Globalization.UmAlQuraCalendar;
//using ThaiBuddhistCalendar = DotNetLib.System.Globalization.ThaiBuddhistCalendar;
//using HijriCalendar = DotNetLib.System.Globalization.HijriCalendar;
//using HebrewCalendar = DotNetLib.System.Globalization.HebrewCalendar;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public static class CalendarExtension
    {
        public static ICalendar Wrap(this GGlobalization.Calendar calendar)
        {
            if (calendar == null) 
            {
                return null;
            }
            if (calendar is GGlobalization.GregorianCalendar gregorianCalendar)
            {
                return new GregorianCalendar(gregorianCalendar);
            }
            if (calendar is GGlobalization.PersianCalendar persianCalendar)
            {
                return new PersianCalendar(persianCalendar);
            }
            if (calendar is GGlobalization.UmAlQuraCalendar umAlQuraCalendar)
            {
                return new UmAlQuraCalendar(umAlQuraCalendar);
            }
            if (calendar is GGlobalization.ThaiBuddhistCalendar thaiBuddhistCalendar)
            {
                return new ThaiBuddhistCalendar(thaiBuddhistCalendar);
            }
            if (calendar is GGlobalization.HijriCalendar hijriCalendar)
            {
                return new HijriCalendar(hijriCalendar);
            }
            if (calendar is GGlobalization.HebrewCalendar hebrewCalendar)
            {
                return new HebrewCalendar(hebrewCalendar);
            }
            if (calendar is GGlobalization.JapaneseCalendar japaneseCalendar)
            {
                return new JapaneseCalendar(japaneseCalendar);
            }
            if (calendar is GGlobalization.KoreanCalendar koreanCalendar)
            {
                return new KoreanCalendar(koreanCalendar);
            }
            if (calendar is GGlobalization.TaiwanCalendar taiwanCalendar)
            {
                return new TaiwanCalendar(taiwanCalendar);
            }
            if (calendar is GGlobalization.ChineseLunisolarCalendar chineseLunisolarCalendar)
            {
                return new ChineseLunisolarCalendar(chineseLunisolarCalendar);
            }
            return null; //If calendar COM wrapper not implemented return null
        }

        public static ICalendar[] Wrap(this GGlobalization.Calendar[] calendars) 
        {
            if (calendars == null)
                return null;

            ICalendar[] wrappedCalendars = new ICalendar[calendars.Length];
            for (int index = 0; index < wrappedCalendars.Length; index++)
            {
                wrappedCalendars[index] = calendars[index].Wrap();
            }
            return wrappedCalendars;
        }

    }
}
