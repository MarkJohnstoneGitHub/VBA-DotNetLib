# VBA DotNetLib COM Interlop
 COM Interlop wrappers of the .Net Framework 4.8.1
 
  Aim: To create .Net Framework 4.8.1 Com Interlop wrappers using C# for VBA 64 to enable various data types in VBA with early and late binding.
 
 Then in VBA create predeclared class wrappers for the DotNetLib COM objects.
 
 Classes initally focussing on are  [DateTime](https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1), [DateTimeOffset](https://learn.microsoft.com/en-us/dotnet/api/system.datetimeoffset?view=netframework-4.8.1) [TimeSpan](https://learn.microsoft.com/en-us/dotnet/api/system.timespan?view=netframework-4.8.1),  [TimeZoneInfo](https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1) and associated classes.
 
 **Dependencies:**
   
   DotNetLib.tlb [DotNetLib.tlb Type library](https://github.com/MarkJohnstoneGitHub/DotNetLib/blob/main/bin/Release/DotNetLib.tlb)
   
   mscorlib.tlb eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
   
 **Usage:**
 
 1) Either building the project in Visual Studio which registers the DotNetLib.tlb or run RegAsm.exe in administrator to register the type library DotNetLib.tlb.
 2) Eg In MS-Access, MS-Excel see Tools->References
 
 Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  
 Add reference mscorlib.tlb
 
 **Issues:**
 
 Currently List COM object wont allow to be created getting invalid use of New Keyword.  This will removed and replaced with it's non-generic equivalent.
 
 Anticipate the ReadOnlyCollection may have similar issue.
 
 Require to consider how to handle generic types in COM Interlop as not supported, possible work around implement each type separately, which enforces type safety.  
 
 Or replace with non-generic equivalent.  To enforce type safety in VBA create a custom wrapper for the collection on the non-generic collection.
 
 **Testing**
 
 Only adhoc testing performed on DateTime and ListString object and appears to create the object and various methods are functional.
 
 **Development Notes**
  
  As COM Interlop doesn't support generic types required to convert to it's non-generic equivalent.
  
  How to treat generic types returned? eg. public static System.Collections.ObjectModel.ReadOnlyCollection<TimeZoneInfo> GetSystemTimeZones()
  
  
  https://github.com/dotnet/platform-compat/blob/master/docs/DE0006.md
 
  https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic?view=netframework-4.8.1
 
