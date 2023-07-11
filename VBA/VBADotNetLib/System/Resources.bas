Attribute VB_Name = "Resources"
'@Folder("VBADotNetLib.System")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12 2023
'@LastModified July 12, 2023

' https://referencesource.microsoft.com/#mscorlib/system/__hresults.cs

Option Explicit


Public Enum COM_HResult
   FormatException = &H80131537                 'COR_E_FORMAT = 0x80131537
   ArgumentOutOfRangeException = &H80131502     'COR_E_ARGUMENTOUTOFRANGE = 0x80131502
   ArgumentNullException = &H80004003           'E_POINTER 0x80004003
   ArgumentException = &H80070057               'COR_E_ARGUMENT = 0x80070057
End Enum
