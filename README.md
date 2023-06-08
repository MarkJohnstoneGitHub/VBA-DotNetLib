# VBA DotNetLib COM Interlop
 COM Interlop wrappers of the .Net Framework 4.8.1
 
  Aim: To create .Net Framework 4.8.1 Com Interlop wrappers using C# for VBA 64 to enable various data types in VBA with early and late binding.
 
 Then in VBA create predeclared class wrappers for the DotNetLib.tlb COM objects.
 
 Classes initally focussing on are  [DateTime](https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1), [DateTimeOffset](https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1) [TimeSpan](https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1),  [TimeZoneInfo](https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1) and associated classes.
 
  **Status:**
  
  Initial development.  Adhoc testing for DateTime only done, DateTimeOffset, TimeZoneInfo, TimeSpan implemented thou not tested, 
  ReadOnlyCollection implemented thou not tested.  TimeZoneInfo.GetSystemTimes returns a ReadOnlyCollection which isn't tested. 
  
  Unit testing will be done once VBA wrappers for COM objects implemented.
  
  Investigating auto generation of VBA COM object wrapper class.
  
  Currently for type library info can obtain all object details expect for class and method description attributes.
  
  Use either Type Library info or C# reflection.  See: [Refactor-COM-object-to-VBA-COM-wrapper-class](https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class)
  

  
 **Dependencies:**
   
[DotNetLib.tlb type library](https://github.com/MarkJohnstoneGitHub/DotNetLib/blob/main/bin/Release/DotNetLib.tlb)
   
mscorlib.tlb type library eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb

If the .NET Framework isn't installed see [Download .NET Framework](https://dotnet.microsoft.com/en-us/download/dotnet-framework)

 **Usage:**
 
 1) Either building the project in Visual Studio which registers the DotNetLib.tlb or run RegAsm.exe in administrator to register the type library DotNetLib.tlb.
 2) Eg In MS-Access, MS-Excel see Tools->References
 
 Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  
 Add reference mscorlib.tlb
 
 The type libraries added can be viewed under View->Object Browser and select DotNetLib 
 
 For detailed explanation of class properties and properties see [netframework-4.8.1](https://learn.microsoft.com/en-us/dotnet/api/system?view=netframework-4.8.1)
 
 **Issues:**
 
 Currently List COM object wont allow to be created getting invalid use of New Keyword.  This will removed and replaced with it's non-generic equivalent.
 
 Anticipate the ReadOnlyCollection may have a similar issue. It has been updated thou not tested.
 
 Require to consider how to handle generic types in COM Interlop as not supported, possible work around implement each type separately, which enforces type safety.  
 
 Or replace with non-generic equivalent.  To enforce type safety in VBA create a custom wrapper for the collection on the non-generic collection.
 
 **Testing**
 
 Only adhoc testing performed on DateTime and ListString object and appears to create the object and various methods are functional.
 
 **Development Notes**
  
  As COM Interlop doesn't support generic types required to convert or wrap to its non-generic equivalent.
  
  How to treat generic types returned? eg. public static System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> GetSystemTimeZones()
  
  
  [DE0006: Non-generic collections shouldn't be used](https://github.com/dotnet/platform-compat/blob/master/docs/DE0006.md)
 
  [System.Collections.Generic Namespace](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic?view=netframework-4.8.1)
 
