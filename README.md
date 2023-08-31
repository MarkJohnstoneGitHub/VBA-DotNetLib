# VBA DotNetLib COM Interop wrappers of the .Net Framework 4.8.1
 
**Aim:** To create .Net Framework 4.8.1 Com Interop wrappers using C# to implement in VBA 64 to enable various .Net Framework data types in VBA with early and/or late binding. Then in VBA create predeclared class wrappers for the DotNetLib.tlb COM objects.  Compatibility intially only VBA 64 on Windows as can only test on windows 64 bit of MS-Office.
 
Classes initally focussing on are [DateTime](https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1), [DateTimeOffset](https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1), [TimeSpan](https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1),  [TimeZoneInfo](https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1) and associated classes.
Added: [CultureInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1), [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1), [NumberFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1) .

 The API for the .Net class may be required to be altered due to VBA reserved words. See [reserved-word-list](https://www.engram9.info/access-2007-vba/reserved-word-list.html).
 
**Status:**

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
  
 **Dependencies:**
 - [DotNetLib.tlb type library](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)
 - mscorlib.tlb type library eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
 - VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3) for the Rubberduck utility exporting components
 - .NET Framework If it is not installed see [Download .NET Framework](https://dotnet.microsoft.com/en-us/download/dotnet-framework)

 **Usage:**
 
 1) Either building the project in Visual Studio which registers the DotNetLib.tlb or run RegAsm.exe in administrator to register the type library [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release).
    - Currently manually installation and registration for type library [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)  See: [register-dll](http://www.geeksengine.com/article/register-dll.html)
    - Copy the [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release) files to a location which don't intend to change eg. C:\ProgramData\DotNetLib then register the DotNetLib type library
    - Eg. To register C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\ProgramData\DotNetLib\DotNetLib.dll /tlb 
    - Eg. To unregister C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\ProgramData\DotNetLib\DotNetLib.dll /tlb /unregister
    - If the files are moved will require to unregister and register manually.
 2) Eg In MS-Access, MS-Excel see Tools->References
   - For [DotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/DotNetLibrary.accdb) references may be required to be fixed by removing and adding back in.
   - Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  i.e. browse to location where stored 
   - Add reference mscorlib.tlb
   - The type libraries added can be viewed under View->Object Browser and select DotNetLib 
 
For detailed explanation of class properties see [netframework-4.8.1](https://learn.microsoft.com/en-us/dotnet/api/system?view=netframework-4.8.1)

Ms Access database [VBADotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/MS-Access) wrapper VBA classes and [examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples) for the DotNetLib.tlb.

 
 **Issues:**
  - [DateTimeFormatInfo.AbbreviatedDayNames](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.abbreviateddaynames?view=netframework-4.8.1)
    - When attempting to assign an array to DateTimeFormatInfo.AbbreviatedDayNames Compile error: Function or interface marked as restricted, or the function uses an Automation typee not supported in Visual Basic
    - https://stackoverflow.com/questions/13185159/how-to-pass-byte-arrays-as-udt-properties-from-vb6-vba-to-c-sharp-com-dll
  - TimeZoneInfo.Local renamed member to Locale.  May cause issues when for interfaces may require renaming in type library? Alternative name?
  - Require to investigate how to correctly marshal arrays 
  - See [PassingParameterArraysByReference](https://www.l3harrisgeospatial.com/docs/PassingParameterArraysByReference.html)
  - [pass-an-array-from-vba-to-c-sharp-using-com-interop](https://stackoverflow.com/questions/2027758/pass-an-array-from-vba-to-c-sharp-using-com-interop)
  - [VBA Interface not showing property in watch window](https://stackoverflow.com/questions/61232755/vba-interface-not-showing-property-in-watch-window)
   -  [how-to-get-property-values-of-classes-that-implement-an-interface-in-the-locals](https://stackoverflow.com/questions/29146243/how-to-get-property-values-of-classes-that-implement-an-interface-in-the-locals)
   -  Work around implement interfaces required in the DotNetLib type library.  They appear to work fine for type library interfaces but not VBA interfaces.
 
Currently List COM object wont allow to be created getting invalid use of New Keyword.  This will removed and replaced with it's non-generic equivalent.
 
Require to consider how to handle generic types in COM Interlop as not supported, possible work around implement each type separately, which enforces type safety.  
 
Or replace with non-generic equivalent.  To enforce type safety in VBA create a custom wrapper for the collection on the non-generic collection.
 
 
 **Development Notes**
  
  As COM Interlop doesn't support generic types required to convert or wrap to its non-generic equivalent.
  
  How to treat generic types returned? eg. public static System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> GetSystemTimeZones()
  
  
  [DE0006: Non-generic collections shouldn't be used](https://github.com/dotnet/platform-compat/blob/master/docs/DE0006.md)
 
  [System.Collections.Generic Namespace](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic?view=netframework-4.8.1)

 
 **Things To do**
 
- Unit testing using Rubberduck unit testing.
- [Create an installer from Microsoft Visual Studio](https://www.advancedinstaller.com/user-guide/tutorial-ai-ext-vs.html#section761)

 Will require implementing the following:
  - [Cultureinfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo?view=netframework-4.8.1) and associated classes. Implemented
  - [Calendar](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.calendar?view=netframework-4.8.1) . Atm considering whether or not to implement.
  - [DateTimeFormat](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1) . Implemented.
  - [CompareInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.compareinfo?view=netframework-4.8.1)
  - [CultureTypes](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.cultureinfo.culturetypes?view=netframework-4.8.1)
  - [NumberFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1) . Implemented thou not tested or currently in use.
  - [DateTimeFormatInfo](https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1). Implemented not fully tested, may effect various DateTime parsing functions.

VBA Wrapper for ReadOnlyCollection for SystemTimeZones i.e. of type TimeZoneInfo

VBA Wrapper for ReadOnlyCollection

Implement interfaces in DotNetLib type library as work around for [VBA Interface not showing property in watch window](https://stackoverflow.com/questions/61232755/vba-interface-not-showing-property-in-watch-window). 
  
