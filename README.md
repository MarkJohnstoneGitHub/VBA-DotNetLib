# VBA DotNetLib COM Interop wrappers of the .Net Framework 4.8.1
 
**Aim:** To create .Net Framework 4.8.1 COM Interop wrappers using C# to implement in VBA 64.  This will enable various .Net Framework data types in VBA with early and/or late binding. Compatibility intially only VBA 64 on Windows as can only test on windows 64 bit of MS-Office. For Mac compatibility would require migrating to .Net Core.
 
Classes initally focussing on are [DateTime](https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1), [DateTimeOffset](https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1), [TimeSpan](https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1),  [TimeZoneInfo](https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1) and associated classes.

Aug 29, 2023 Added: [CultureInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1), [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1), [NumberFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1), [TextInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1) .

Sep 19, 2023 Added: [ChineseLunisolarCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1),  [GregorianCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.gregoriancalendar?view=netframework-4.8.1), [HebrewCalendar ](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hebrewcalendar?view=netframework-4.8.1), [HijriCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hijricalendar?view=netframework-4.8.1), [JapaneseCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.japanesecalendar?view=netframework-4.8.1), [JulianCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.juliancalendar?view=netframework-4.8.1), [KoreanCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.koreancalendar?view=netframework-4.8.1), [PersianCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.persiancalendar?view=netframework-4.8.1), [ThaiBuddhistCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.thaibuddhistcalendar?view=netframework-4.8.1), [UmAlQuraCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.umalquracalendar?view=netframework-4.8.1)

Sep 22, 2023 Added: [CompareInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1) 

Sep 23, 2023 Added [String](https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1) For the VBA singleton wrapper renamed String to Strings due to VBA reserved word.

Sep 25, 2023 Added [Regex](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1) Implemented so far Regex.Unescape and Regex.Escape

Sep 30, 2023 Added [System.Text.RegularExpressions](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions?view=netframework-4.8.1), [Capture](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capture?view=netframework-4.8.1), [CaptureCollection](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capturecollection?view=netframework-4.8.1), [Group](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1), [GroupCollection](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.groupcollection?view=netframework-4.8.1), [Match](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1), [MatchCollection](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchcollection?view=netframework-4.8.1), [Regex](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1)

Oct 3, 2023 Added [ListString](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1?view=netframework-4.8.1) 
- Attempted to dynamically create a List providing the data type instance however having issues casting, therefore will wrap a List for various basic types individually.
- Testing still to be done. Create, Add, BinarySearch, Contains, IndexOf, Insert, Reverse, Sort, appears functioning correctly.

Oct 5, 2023 Added [ArrayList](https://learn.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8.1)

Oct 12, 2023 Added [Array](https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1), [Type](https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1), [GenericParameterAttributes](https://learn.microsoft.com/en-us/dotnet/api/system.reflection.genericparameterattributes?view=netframework-4.8.1)

Oct 15, 2023 Added [Queue](https://learn.microsoft.com/en-us/dotnet/api/system.collections.queue?view=netframework-4.8.1), [Stack](https://learn.microsoft.com/en-us/dotnet/api/system.collections.stack?view=netframework-4.8.1)

Oct 16, 2023 Added [SortedList](https://learn.microsoft.com/en-us/dotnet/api/system.collections.sortedlist?view=netframework-4.8.1)

Oct 17, 2023 Added [CaseInsensitiveComparer](https://learn.microsoft.com/en-us/dotnet/api/system.collections.caseinsensitivecomparer?view=netframework-4.8.1), [StringComparer](https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparer?view=netframework-4.8.1)

Oct 18, 2023 Added [DictionaryEntry](https://learn.microsoft.com/en-us/dotnet/api/system.collections.dictionaryentry?view=netframework-4.8.1), [Hashtable](https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1)

Oct 31, 2023 Added [StringBuilder](https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder?view=netframework-4.8.1)

Nov 2, 2023 Added [BitArray](https://learn.microsoft.com/en-us/dotnet/api/system.collections.bitarray?view=netframework-4.8.1)

Nov 16, 2023 Added [Directory](https://learn.microsoft.com/en-us/dotnet/api/system.io.directory?view=netframework-4.8.1), [DirectoryInfo](https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1), [Environment](https://learn.microsoft.com/en-us/dotnet/api/system.environment?view=netframework-4.8.1), [File](https://learn.microsoft.com/en-us/dotnet/api/system.io.file?view=netframework-4.8.1),  [FileInfo](https://learn.microsoft.com/en-us/dotnet/api/system.io.fileinfo?view=netframework-4.8.1), [Path](https://learn.microsoft.com/en-us/dotnet/api/system.io.path?view=netframework-4.8.1) , [FileSystemInfo](https://learn.microsoft.com/en-us/dotnet/api/system.io.filesysteminfo?view=netframework-4.8.1), [StreamWriter](https://learn.microsoft.com/en-us/dotnet/api/system.io.streamwriter?view=netframework-4.8.1) [AccessControlSections](https://learn.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.accesscontrolsections?view=netframework-4.8.1), [SpecialFolderOption](https://learn.microsoft.com/en-us/dotnet/api/system.environment.specialfolderoption?view=netframework-4.8.1), [SpecialFolders](https://learn.microsoft.com/en-us/dotnet/api/system.environment.specialfolder?view=netframework-4.8.1), [StringSplitOptions](https://learn.microsoft.com/en-us/dotnet/api/system.stringsplitoptions?view=netframework-4.8.1)

Nov 23, 2023 Added [ASCIIEncoding](https://learn.microsoft.com/en-us/dotnet/api/system.text.asciiencoding?view=netframework-4.8.1), [Encoding](https://learn.microsoft.com/en-us/dotnet/api/system.text.encoding?view=netframework-4.8.1), [UnicodeEncoding](https://learn.microsoft.com/en-us/dotnet/api/system.text.unicodeencoding?view=netframework-4.8.1), [UTF32Encoding](https://learn.microsoft.com/en-us/dotnet/api/system.text.utf32encoding?view=netframework-4.8.1), [UTF7Encoding](https://learn.microsoft.com/en-us/dotnet/api/system.text.utf7encoding?view=netframework-4.8.1), [UTF8Encoding](https://learn.microsoft.com/en-us/dotnet/api/system.text.utf8encoding?view=netframework-4.8.1)

 **Affected API due to VBA reserved words:**

 The API for the .Net class or VBA singletons for associated .Net classes may be required to be altered due to VBA reserved words. See [reserved-word-list](https://www.engram9.info/access-2007-vba/reserved-word-list.html).
 
  - [TimeZoneInfo.Local](https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.local?view=netframework-4.8.1) renamed to TimeZoneInfo.Locale.
  - [String](https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1) VBA singleton renamed to Strings.
  - [Array](https://learn.microsoft.com/en-us/dotnet/api/system.array?view=netframework-4.8.1) VBA singleton renamed to Arrays.
  - [Type](https://learn.microsoft.com/en-us/dotnet/api/system.type?view=netframework-4.8.1) VBA singleton renamed to Types.

As VBA doesnot have member overloading factory methods and member overloads will differ.  Overloads generally are named with a preceeding number. Unique naming maybe used for factory methods.

 **Dependencies:**
 - [DotNetLib.tlb type library](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)
 - mscorlib.tlb type library eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
 - VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3) for the Rubberduck utility exporting components
 - .NET Framework If it is not installed see [Download .NET Framework](https://dotnet.microsoft.com/en-us/download/dotnet-framework)

 **Usage:**
 
 1) Register [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)
    - Either building the project in Visual Studio which registers the DotNetLib.tlb or run RegAsm.exe in administrator to register the type library [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release).
    - Currently manually installation and registration for type library [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)  See: [register-dll](http://www.geeksengine.com/article/register-dll.html)
    - Copy the [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release) files to a location which don't intend to change eg. C:\ProgramData\DotNetLib then register the DotNetLib type library
    - Eg. To register C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\ProgramData\DotNetLib\DotNetLib.dll /tlb 
    - Eg. To unregister C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\ProgramData\DotNetLib\DotNetLib.dll /tlb /unregister
    - If the files are moved will require to unregister and register manually.
    - If the DotNetLib type library is updated will require to unregister and register manually.
 2) Add References required.
    - Eg In MS-Access, MS-Excel see Tools->References
    - For [DotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/DotNetLibrary.accdb) references may be required to be fixed by removing and adding back in.
    - Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  i.e. browse to location where stored 
    - Add reference mscorlib.tlb version 2.4
    - Add reference Microsoft VBScript Regular Expressions 5.3 (Required only for the [Rubberduck export utility](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Rubberduck%20Utility) not part of VBADotNetLib)
    - The type libraries added can be viewed under View->Object Browser and select DotNetLib.tlb
3) Add the [VBADotNetLib](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/VBADotNetLib/) VBA Factory/Singleton classes into a project.
    - Either copy the classes or add a reference to project containing the classes.
4) Recommended install [Rubberduck](https://rubberduckvba.com/) VBA Addin.
 
For detailed explanation of the DotNetLib class properties see [netframework-4.8.1](https://learn.microsoft.com/en-us/dotnet/api/system?view=netframework-4.8.1)

Ms Access database [VBADotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/MS-Access) VBA Factory classes and [examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples) for the DotNetLib.tlb. Also a MS-Excel version [VBADotNetLib.xlsm](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Ms-Excel) .  

Note: The MS-Access contains the latest version of VBADotNetLibrary and examples as the development is performed in MS-Access and periodically exported to the VBADotNetLibrary MS-Excel spreadsheet. 

 
 **Regular expressions : Converting strings containing escape sequences and special characterss:**
 
 To use escape and special characters. Note if require quotes " require to escape in VBA with double quotes.

VBA Example using Regex.Unescape with hexadecimal escape sequences
```
    Dim stringUpper As String
    stringUpper = "\x41\x42\x43"     ' Create upper-case characters from their Unicode code units.
    stringUpper = Regex.Unescape(stringUpper)
    Debug.Print stringUpper
    'Output: ABC
```

 **Issues:**
 
Hashtable.Item(key) = valuetype causes an Object required error for value types.  Added member SetValue(key,value) to use as an alternative until fixed.
- To fix requires creating an IDL and manually adding a propput for value types and compiling type library with MIDL.

ArrayList.Item(index) = item
- Cannot assign value types using arraylist.Item(index) = valueType
- Eg. ```pvtStringList.Item(2) = "abcd" ``` Will produce a Run-time error 424 Object required
- To cater for value types added the Arraylist.SetItem(index,item) member.  Arraylist.SetItem(index,item) can be use for value or object types.
- Eg. assigning a value type ```pvtStringList.SetItem 2, "abcd"```

 - Currently List COM object wont allow to be created getting invalid use of New Keyword. This will removed and replaced with it's non-generic equivalent.. 
- Too many things to do. Argh!
 
 **Things To do**
 
- Unit testing using Rubberduck unit testing.
- [Create an installer from Microsoft Visual Studio](https://www.advancedinstaller.com/user-guide/tutorial-ai-ext-vs.html#section761)


**Update History**

**Status: Latest Updates**

**DotNetLib Update September 30th, 2023** 

Added [System.Text.RegularExpressions](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions?view=netframework-4.8.1), [Capture](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capture?view=netframework-4.8.1), [CaptureCollection](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.capturecollection?view=netframework-4.8.1), [Group](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.group?view=netframework-4.8.1), [GroupCollection](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.groupcollection?view=netframework-4.8.1), [Match](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.match?view=netframework-4.8.1), [MatchCollection](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.matchcollection?view=netframework-4.8.1), [Regex](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1)

Todo:
- Implement VBA singleton classes for Match and Group for static members.
- Examples and unit testing.


**DotNetLib Update September 25th, 2023** 
Added [Regex](https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex?view=netframework-4.8.1) 
- Implemented Regex.Unescape and Regex.Escape
- Regex.Unescape can be used to convert VBA literal strings containing escape characters.

Updated Strings, added the following members
- Compare, CompareOrdinal, Copy, Equals, IsNullOrEmpty, IsNullOrWhiteSpace

**DotNetLib Update September 23rd, 2023** 

Added [String](https://learn.microsoft.com/en-us/dotnet/api/system.string?view=netframework-4.8.1) 
 - So far only implemented static members [String.Format](https://learn.microsoft.com/en-us/dotnet/api/system.string.format?view=netframework-4.8.1)
 - Renamed String to Strings due to VBA reserved word.

Added: IFormatProviderExtension.cs to UnWrap IFormatProvider types.

**DotNetLib Update September 22nd, 2023** 
- Renamed abstract class ICalendar to Calendar to keep consistent with Net Framework
- Updated VBADotNetLib for affected calendar classes and examples.
- Added CompareInfo, CultureInfo.CompareInfo member properties now availble.
- Todo add to VBADotNetLib CompareInfo singleton class.

**DotNetLib Update September 20th, 2023** 

Updated DateTime.cs, IDateTime.cs, DateTimeSingleton, IDateTimeSingleton.cs, 
- Added factory methods for ICalendar parameter.
- public DateTime CreateFromDate2(int pYear, int pMonth, int pDay, ICalendar calendar)
- public DateTime CreateFromDateTime2(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, ICalendar calendar)
- public DateTime CreateFromDateTime3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, ICalendar calendar)
- public DateTime CreateFromDateTimeKind3(int pYear, int pMonth, int pDay, int pHour, int pMinute, int pSecond, int pMillisecond, ICalendar calendar, DateTimeKind pKind)
- [DateTime.cls](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/DateTime.cls) added the above new factory methods available from DateTimeSingleton DotNetLib.tlb.
 
Todo add examples and testing.
- Update DotNetLib class members that reference the Calendar class.
- Eg. DateTime constructors, DateTimeOffset constructors
- Updated DateTimeFormatInfo.Calendar member to use wrapped ICalendar.  DateTimeFormatInfo.Calendar property should now be available to access and set. (Require to test)

**DotNetLib Update September 19th, 2023** 
- Implemented abstract class [Calendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendar?view=netframework-4.8.1) as ICalendar, updated CultureInfo for properties Calendar and OptionalCalendars which are now availbable and added the following calendars:
   - [ChineseLunisolarCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.chineselunisolarcalendar?view=netframework-4.8.1)
   - [GregorianCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.gregoriancalendar?view=netframework-4.8.1)
   - [HebrewCalendar ](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hebrewcalendar?view=netframework-4.8.1)
   - [HijriCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.hijricalendar?view=netframework-4.8.1)
   - [JapaneseCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.japanesecalendar?view=netframework-4.8.1)
   - [JulianCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.juliancalendar?view=netframework-4.8.1)
   - [KoreanCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.koreancalendar?view=netframework-4.8.1)
   - [PersianCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.persiancalendar?view=netframework-4.8.1)
   - [TaiwanCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.taiwancalendar?view=netframework-4.8.1)
   - [ThaiBuddhistCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.thaibuddhistcalendar?view=netframework-4.8.1)
   - [UmAlQuraCalendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.umalquracalendar?view=netframework-4.8.1)

 Todo testing for creating each added calendar, CultureInfo.Calendar, CultureInfo.OptionalCalendars.  Adhoc testing not detecting any missing Calendars required for the default Calendar or optional calendars.
 - Update [DateTimeFormatInfo.Calendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.calendar?view=netframework-4.8.1) to use  ICalendar

**DotNetLib Update September 9th, 2023** 
- API changes for DateTime, DateTimeOffset, TimeSpan
   - Merged member ToString4(string format, IFormatProvider provider) and replace with ToString2(string format, IFormatProvider provider = null)
   - Updated examples using ToString4(string format, IFormatProvider provider) to  use ToString2(string format, IFormatProvider provider = null) due to DotNetLib.tlb API changes.
   - Add  IComparable, IFormattable interfaces
- Added Console.cls Not fully functional (Work in progress)

**DotNetLib Update September 5th, 2023** 
 - Added [TextInfo Class](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1). Properties for [CultureInfo.TextInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.textinfo?view=netframework-4.8.1) now available.
 - Added VBA wrapper TextInfo singleton class for TextInfo.ReadOnly(TextInfo) method.
 - Todo update [Ms-Excel VBADotNetLib.xlsm](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/Ms-Excel/VBADotNetLib.xlsm)

**DotNetLib Update September 3rd, 2023** 
 - For DateTime and DateTimeOffset renamed DateOnly property to Date property to be consistent with .Net  Framework. 
 - Updated all effected examples

**DotNetLib Update September 2nd, 2023** 
 - Fixed issues with [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1)
   - Changed format parameter to string from char
   - public string[] GetAllDateTimePatterns(string format = null)
   - public void SetAllDateTimePatterns([In] ref string[] patterns, string format)
- Added overloads for [DateTime.GetDateTimeFormats](https://learn.microsoft.com/en-us/dotnet/api/system.datetime.getdatetimeformats?view=netframework-4.8.1)
   - Changed format parameter to string from char
- Refactored [CultureInfo.cls](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/Globalization/CultureInfo.cls) renamed constructors to more meaningful names and combined overloads using an optional parameter.
     
**DotNetLib Update September 1st, 2023** 
 - Fixed issues with [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1)
 - When assigning an array to a property eg DateTimeFormatInfo.AbbreviatedDayNames Compile error: Function or interface marked as restricted, or the function uses an Automation type not supported in Visual Basic
 - https://stackoverflow.com/questions/13185159/how-to-pass-byte-arrays-as-udt-properties-from-vb6-vba-to-c-sharp-com-dll
 - Added members to set the various arrays replacing the set propeterty which is no longer COM visibile.
   - SetAbbreviatedDayNames([In] ref string[] abbreviatedDayNames)
   - SetAbbreviatedMonthGenitiveNames([In] ref string[] abbreviatedMonthGenitiveNames)
   - SetDayNames([In] ref string[] dayNames)
   - SetMonthGenitiveNames([In] ref string[] monthGenitiveNames)
   - SetMonthNames([In] ref string[] monthNames)
   - SetShortestDayNames([In] ref string[] shortestDayNames)

**DotNetLib Update August 29th, 2023** 

 - Added [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1)
 - Added [NumberFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1)
 - DateTimeFormatInfo and NumberFormatInfo properties are now available for [Cultureinfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1)
 - Unit Testing required to test that various DateTime, DateTimeOffset, TimeSpan parsing functions using [IFormatProvider](https://learn.microsoft.com/en-us/dotnet/api/system.iformatprovider?view=netframework-4.8.1) is functioning correctly. Adhoc testing done using examples.

**DotNetLib Version 1.2 Update August 17th, 2023** 

Completed rewritting the DotNetLib type library and VBA DotNetLib wrappers to use the [Singleton pattern](https://en.wikipedia.org/wiki/Singleton_pattern).
Where factory methods and static members are in a singleton classs.

Currently the default interfaces IDateTime, IDateTimeOffset, ITimeSpan, ITimeZoneInfo, ICultureInfo for its corresponding COM object isn't displayed in the VBA Object browser or editor thou accessible. Can program either directly against the COM Object eg. ```Dim myDateTime as DotNetLib.DateTime``` or its interface ```Dim myDateTime as IDateTime``` 

For the creation and access of static members use its corresponding Singleton/Factory class eg ```Set myDateTime = DateTime.CreateFromDate(2010, 8, 18) ```

**Initial developement.**
 - API of the type library and VBA COM wrapper classes may be altered during initial development.
 - Implemented the following C# COM Interlop wrappers of the .Net Framework 4.8.1 [DotNetLib type library](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib), see [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)
     - [DateTime](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/DateTime.cs)
     - [DateTimeOffset](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/DateTimeOffset.cs)
     - [TimeSpan](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/TimeSpan.cs)
     - [TimeZoneInfo](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/TimeZoneInfo.cs)
     - [CultureInfo](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/Globalization/CultureInfo.cls)
     - [DateTimeKind enum](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/DateTimeKind.cs)
     - [DayOfWeek enum ](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/DayOfWeek.cs)
     - [TimeSpanStyles enum](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/Globalization/TimeSpanStyles.cs)
     - [ReadOnlyCollection](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/COMDotNetLib/System/Collections/ReadOnlyCollection.cs)
     - Adhoc testing using VBA examples located in [DotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/DotNetLibrary.accdb)
- VBA [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release) COM Wrappers implemented.
  - [CultureInfo](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/Globalization/CultureInfo.cls)
  - [DateTime](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/DateTime.cls) adhoc testing and [DateTime examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples/DateTime).
  - [DateTimeOffset](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/DateTimeOffset.cls) adhoc testing and [DateTimeOffset examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples/DateTimeOffset).
  - [TimeSpan](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/TimeSpan.cls) adhoc testing and [TimeSpan examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples/TimeSpan).
  - [TimeZoneInfo](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/VBADotNetLib/System/TimeZoneInfo.cls) adhoc testing and [TimeZoneInfo examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples/TimeZoneInfo)
  - ReadOnlyCollection VBA wrapper to be implemented. DotNetLib.ReadOnlyCollection adhoc testing with TimeZoneInfo examples. TimeZoneInfo.GetSystemTimeZones returns a ReadOnlyCollection. 
  - Unit testing aim to do once VBA wrappers for COM objects implemented.
  - Investigated auto generation of VBA COM object wrapper class. See: [Refactor-COM-object-to-VBA-COM-wrapper-class](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class)

VBA Wrapper for ReadOnlyCollection

Implement interfaces in DotNetLib type library as work around for [VBA Interface not showing property in watch window](https://stackoverflow.com/questions/61232755/vba-interface-not-showing-property-in-watch-window). 

 **Development Notes**
  
  As COM Interlop doesn't support generic types required to convert or wrap to its non-generic equivalent.
  
  How to treat generic types returned? eg. public static System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> GetSystemTimeZones()
  
  
  [DE0006: Non-generic collections shouldn't be used](https://github.com/dotnet/platform-compat/blob/master/docs/DE0006.md)
 
  [System.Collections.Generic Namespace](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic?view=netframework-4.8.1)

Require to investigate how to correctly marshal arrays 
  - See [PassingParameterArraysByReference](https://www.l3harrisgeospatial.com/docs/PassingParameterArraysByReference.html)
  - [pass-an-array-from-vba-to-c-sharp-using-com-interop](https://stackoverflow.com/questions/2027758/pass-an-array-from-vba-to-c-sharp-using-com-interop)

      - [DateTimeFormatInfo.AbbreviatedDayNames](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.abbreviateddaynames?view=netframework-4.8.1)
    - When attempting to assign an array to DateTimeFormatInfo.AbbreviatedDayNames Compile error: Function or interface marked as restricted, or the function uses an Automation typee not supported in Visual Basic
    - https://stackoverflow.com/questions/13185159/how-to-pass-byte-arrays-as-udt-properties-from-vb6-vba-to-c-sharp-com-dll
    - Fixed by implementing set methods and making set property not COM visible.
  - TimeZoneInfo.Local renamed member to Locale. 
  - [VBA Interface not showing property in watch window](https://stackoverflow.com/questions/61232755/vba-interface-not-showing-property-in-watch-window)
   -  [how-to-get-property-values-of-classes-that-implement-an-interface-in-the-locals](https://stackoverflow.com/questions/29146243/how-to-get-property-values-of-classes-that-implement-an-interface-in-the-locals)
   -  Work around implement interfaces required in the DotNetLib type library.  They appear to work fine for type library interfaces but not VBA interfaces.
 
Currently List COM object wont allow to be created getting invalid use of New Keyword.  This will removed and replaced with it's non-generic equivalent.

 Will require implementing the following:
  - [Cultureinfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1) and associated classes. Implemented
   - [Calendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendar?view=netframework-4.8.1) . Implemented, currently updating class members referencing the Calendar class.
   - [DateTimeFormat](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1) . Implemented.
   - [CompareInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1)
   - [CultureTypes](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.culturetypes?view=netframework-4.8.1)
   - [NumberFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1) . Implemented thou not tested or currently in use.
   - [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1). Implemented not fully tested, may effect various DateTime parsing functions.
   - [TextInfo Class](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1)

VBA Wrapper for ReadOnlyCollection for SystemTimeZones i.e. of type TimeZoneInfo
 
Require to consider how to handle generic types in COM Interlop as not supported, possible work around implement each type separately, which enforces type safety.  
 
Or replace with non-generic equivalent.  To enforce type safety in VBA create a custom wrapper for the collection on the non-generic collection.

 **Collections List**

How to create dynamic list? I.e. When creating a List  specify the type required.  
- https://stackoverflow.com/questions/9860387/how-do-i-create-a-dynamic-type-listt
  
