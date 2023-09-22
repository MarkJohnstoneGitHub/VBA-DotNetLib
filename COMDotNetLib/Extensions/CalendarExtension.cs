using DotNetLib.System.Globalization;
using GGlobalization = global::System.Globalization;
using System.Runtime.InteropServices;

namespace DotNetLib.Extensions
{
    [ComVisible(false)]
    public static class CalendarExtension
    {
        public static Calendar Wrap(this GGlobalization.Calendar pCalendar)
        {
            if (pCalendar == null) 
            {
                return null;
            }
            if (pCalendar is GGlobalization.GregorianCalendar gregorianCalendar)
            {
                return new GregorianCalendar(gregorianCalendar);
            }
            if (pCalendar is GGlobalization.PersianCalendar persianCalendar)
            {
                return new PersianCalendar(persianCalendar);
            }
            if (pCalendar is GGlobalization.UmAlQuraCalendar umAlQuraCalendar)
            {
                return new UmAlQuraCalendar(umAlQuraCalendar);
            }
            if (pCalendar is GGlobalization.ThaiBuddhistCalendar thaiBuddhistCalendar)
            {
                return new ThaiBuddhistCalendar(thaiBuddhistCalendar);
            }
            if (pCalendar is GGlobalization.HijriCalendar hijriCalendar)
            {
                return new HijriCalendar(hijriCalendar);
            }
            if (pCalendar is GGlobalization.HebrewCalendar hebrewCalendar)
            {
                return new HebrewCalendar(hebrewCalendar);
            }
            if (pCalendar is GGlobalization.JapaneseCalendar japaneseCalendar)
            {
                return new JapaneseCalendar(japaneseCalendar);
            }
            if (pCalendar is GGlobalization.KoreanCalendar koreanCalendar)
            {
                return new KoreanCalendar(koreanCalendar);
            }
            if (pCalendar is GGlobalization.TaiwanCalendar taiwanCalendar)
            {
                return new TaiwanCalendar(taiwanCalendar);
            }
            if (pCalendar is GGlobalization.ChineseLunisolarCalendar chineseLunisolarCalendar)
            {
                return new ChineseLunisolarCalendar(chineseLunisolarCalendar);
            }
            return null; //If pCalendar COM wrapper not implemented return null
        }

        public static Calendar[] Wrap(this GGlobalization.Calendar[] calendars) 
        {
            if (calendars == null)
                return null;

            Calendar[] wrappedCalendars = new Calendar[calendars.Length];
            for (int index = 0; index < wrappedCalendars.Length; index++)
            {
                wrappedCalendars[index] = calendars[index].Wrap();
            }
            return wrappedCalendars;
        }

    }
}
