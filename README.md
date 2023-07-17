# VBA DotNetLib COM Interlop
 COM Interlop wrappers of the .Net Framework 4.8.1
 
  Aim: To create .Net Framework 4.8.1 Com Interlop wrappers using C# to implement in VBA 64 to enable various .Net Framework data types in VBA with early and/or late binding. Then in VBA create predeclared class wrappers for the DotNetLib.tlb COM objects.
 
 Classes initally focussing on are  [DateTime](https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1), [DateTimeOffset](https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1), [TimeSpan](https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1),  [TimeZoneInfo](https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1) and associated classes.
 
  **Status:**
  
  Initial development.  
  - Adhoc testing and [DateTime examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples/DateTime).
  - Adhoc testing and [TimeSpan examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples/TimeSpan). 
  - DateTimeOffset, TimeZoneInfo implemented thou not tested.
  - ReadOnlyCollection implemented thou not tested.  TimeZoneInfo.GetSystemTimes returns a ReadOnlyCollection which isn't tested. 
  - Unit testing aim to do once VBA wrappers for COM objects implemented.
  - Investigated auto generation of VBA COM object wrapper class. See: [Refactor-COM-object-to-VBA-COM-wrapper-class](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class)
  
 **Dependencies:**
 - [DotNetLib.tlb type library](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)
 - mscorlib.tlb type library eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
 - .NET Framework If it is not installed see [Download .NET Framework](https://dotnet.microsoft.com/en-us/download/dotnet-framework)

 **Usage:**
 
 1) Either building the project in Visual Studio which registers the DotNetLib.tlb or run RegAsm.exe in administrator to register the type library [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release).
    - Currently manually installation and registration for type library [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release)
    - Copy the [DotNetLib.tlb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/COMDotNetLib/bin/Release) files to a location which don't intend to change eg. C:\ProgramData\DotNetLib then register the DotNetLib type library
    - Eg. To register C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\ProgramData\DotNetLib\DotNetLib.dll /tlb 
    - Eg. To unregister C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm.exe C:\ProgramData\DotNetLib\DotNetLib.dll /tlb /unregister
    - If the files are moved will require to unregister and register manually.
 2) Eg In MS-Access, MS-Excel see Tools->References
   - For [DotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/DotNetLibrary.accdb) references may be required to be fixed by removing and adding back in.
   - Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  i.e. browse to location where stored 
   - Add reference mscorlib.tlb
   - The type libraries added can be viewed under View->Object Browser and select DotNetLib 
 
For detailed explanation of class properties and properties see [netframework-4.8.1](https://learn.microsoft.com/en-us/dotnet/api/system?view=netframework-4.8.1)

Ms Access database [DotNetLibrary.accdb](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/blob/main/VBA/DotNetLibrary.accdb) wrapper VBA classes and [examples](https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib/tree/main/VBA/Examples) for the DotNetLib.tlb.

 
 **Issues:**
  - Require to investigate how to correctly marshal arrays
  - See [PassingParameterArraysByReference](https://www.l3harrisgeospatial.com/docs/PassingParameterArraysByReference.html)
  - [pass-an-array-from-vba-to-c-sharp-using-com-interop](https://stackoverflow.com/questions/2027758/pass-an-array-from-vba-to-c-sharp-using-com-interop)
 
 Currently List COM object wont allow to be created getting invalid use of New Keyword.  This will removed and replaced with it's non-generic equivalent.
 
 Anticipate the ReadOnlyCollection may have a similar issue. It has been updated thou not tested.
 
 Require to consider how to handle generic types in COM Interlop as not supported, possible work around implement each type separately, which enforces type safety.  
 
 Or replace with non-generic equivalent.  To enforce type safety in VBA create a custom wrapper for the collection on the non-generic collection.
 
 
 **Development Notes**
  
  As COM Interlop doesn't support generic types required to convert or wrap to its non-generic equivalent.
  
  How to treat generic types returned? eg. public static System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> GetSystemTimeZones()
  
  
  [DE0006: Non-generic collections shouldn't be used](https://github.com/dotnet/platform-compat/blob/master/docs/DE0006.md)
 
  [System.Collections.Generic Namespace](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic?view=netframework-4.8.1)
 
