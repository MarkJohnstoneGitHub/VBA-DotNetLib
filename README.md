# VBA DotNetLib COM Interlop
 COM Interlop wrappers of the .Net Framework 4.8.1
 
   **Dependencies:**
   
   DotNetLib.tlb https://github.com/MarkJohnstoneGitHub/DotNetLib/blob/main/bin/Debug/DotNetLib.tlb
   
   mscorlib.tlb eg Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
   
 
 Aim: To create .Net Framework 4.8.1 Com Interlop wrappers using C# for VBA 64.
 
 Then in VBA create predeclared class wrappers for the DotLib COM objects.
 
 Classes initally focussing on are DateTime,TimeSpan,TimeZoneInfo and associated classes.
 
 **Usage:**
 In VBA see Tools->References
 
 Add reference DotNetlib.tlb (Com Interlop wrappers of the .Net Framework 4.8.1)  
 Add reference mscorlib.tlb
 
 
 Issues:
 Currently List COM object wont allow to be created getting invalid use of New Keyword.
   
