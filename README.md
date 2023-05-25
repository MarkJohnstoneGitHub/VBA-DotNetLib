# VBA DotNetLib COM Interlop
 COM Interlop wrappers of the .Net Framework 4.8.1
 
   **Dependencies:**
   
   DotNetLib.tlb https://github.com/MarkJohnstoneGitHub/DotNetLib/blob/main/bin/Debug/DotNetLib.tlb
   
   mscorlib.tlb eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
   
 
 Aim: To create .Net Framework 4.8.1 Com Interlop wrappers using C# for VBA 64.
 
 Then in VBA create predeclared class wrappers for the DotLib COM objects.
 
 Classes initally focussing on are DateTime,TimeSpan,TimeZoneInfo and associated classes.
 
 **Usage:**
 
 1) Either building the project in Visual Studio which registers the DotNetLib.tlb or run RegAsm.exe in administrator to register the type library DotNetLib.tlb.
 2) In VBA see Tools->References
 
 Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  
 Add reference mscorlib.tlb
 
  
 
 **Issues:**
 
 Currently List COM object wont allow to be created getting invalid use of New Keyword.
 
 Anticipate the ReadOnlyCollection may have similar issue.
 
 Require to consider how to handle generic types in COM Interlop as not supported, possible work around implement each type separately, which enforces type safety.  
 
 **Testing**
 
 Only adhoc testing performed on DateTime object and appears to create the object and various methods are functional.
