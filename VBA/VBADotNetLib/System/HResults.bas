Attribute VB_Name = "HResults"
'@Folder("VBADotNetLib.System")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12 2023
'@LastModified July 12, 2023

' https://referencesource.microsoft.com/#mscorlib/system/__hresults.cs

Option Explicit

Const COR_E_FORMAT As Long = &H80131537
Const COR_E_ARGUMENTOUTOFRANGE As Long = &H80131502
Const E_POINTER As Long = &H80004003
Const COR_E_ARGUMENT As Long = &H80070057

Public Enum COMHResult
   FormatException = COR_E_FORMAT
   ArgumentOutOfRangeException = COR_E_ARGUMENTOUTOFRANGE
   ArgumentNullException = E_POINTER
   ArgumentException = COR_E_ARGUMENT
End Enum
