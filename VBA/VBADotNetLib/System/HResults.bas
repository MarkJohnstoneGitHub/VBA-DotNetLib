Attribute VB_Name = "HResults"
'@IgnoreModule ConstantNotUsed
'@Folder "VBADotNetLib.System"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12, 2023
'@LastModified January 28, 2024

' https://referencesource.microsoft.com/#mscorlib/system/__hresults.cs

' https://powershellexplained.com/2017-04-07-all-dotnet-exception-list/

Option Explicit

Const E_FAIL                        As Long = &H80004005
Const E_POINTER                     As Long = &H80004003
Const COR_E_NULLREFERENCE           As Long = &H80004003
Const E_NOTIMPL                     As Long = &H80004001
Const COR_E_FORMAT                  As Long = &H80131537
Const COR_E_ARGUMENTOUTOFRANGE      As Long = &H80131502
Const COR_E_ARGUMENT                As Long = &H80070057
Const COR_E_OVERFLOW                As Long = &H80131516
Const COR_E_EXCEPTION               As Long = &H80131500
Const COR_E_OUTOFMEMORY             As Long = &H8007000E
Const COR_E_INVALIDOPERATION        As Long = &H80131509
Const COR_E_INVALIDCAST             As Long = &H80004002
Const COR_E_NOTSUPPORTED            As Long = &H80131515
Const COR_E_ARRAYTYPEMISMATCH       As Long = &H80131503
Const COR_E_TARGETINVOCATION        As Long = &H80131604
Const COR_E_TYPELOAD                As Long = &H80131522
Const COR_E_BADIMAGEFORMAT          As Long = &H8007000B
Const COR_E_INDEXOUTOFRANGE         As Long = &H80131508
Const COR_E_UNAUTHORIZEDACCESS      As Long = &H80070005
Const COR_E_PLATFORMNOTSUPPORTED    As Long = &H80131539
Const COR_E_SECURITY                As Long = &H8013150A

' https://referencesource.microsoft.com/#mscorlib/system/io/__hresults.cs
Const COR_E_ENDOFSTREAM             As Long = &H80070026
Const COR_E_FILELOAD                As Long = &H80131621
Const COR_E_FILENOTFOUND            As Long = &H80070002
Const COR_E_DIRECTORYNOTFOUND       As Long = &H80070003    '-2147024893
Const COR_E_PATHTOOLONG             As Long = &H800700CE
Const COR_E_IO                      As Long = &H80131620

Public Enum COMHResult
    ArgumentOutOfRangeException = COR_E_ARGUMENTOUTOFRANGE
    ArgumentNullException = E_POINTER
    ArgumentException = COR_E_ARGUMENT
    FormatException = COR_E_FORMAT
    NotImplementedException = E_NOTIMPL
    OutOfMemoryException = COR_E_OUTOFMEMORY
    OverflowException = COR_E_OVERFLOW
    TimeZoneNotFoundException = COR_E_EXCEPTION
    InvalidTimeZoneException = COR_E_EXCEPTION '@TODO Check
    CultureNotFoundException = COR_E_ARGUMENT
    InvalidOperationException = COR_E_INVALIDOPERATION
    InvalidCastException = COR_E_INVALIDCAST
    NotSupportedException = COR_E_NOTSUPPORTED
    TargetInvocationException = COR_E_TARGETINVOCATION
    TypeLoadException = COR_E_TYPELOAD
    FileLoadException = COR_E_FILELOAD      ' Default HRESULT of COR_E_FILELOAD, which has the value 0x80131621, but this is not the only possible HRESULT.
    BadImageFormatException = COR_E_BADIMAGEFORMAT
    IndexOutOfRangeException = COR_E_INDEXOUTOFRANGE
    IOException = COR_E_IO
    UnauthorizedAccessException = COR_E_UNAUTHORIZEDACCESS
    PathTooLongException = COR_E_PATHTOOLONG
    DirectoryNotFoundException = COR_E_DIRECTORYNOTFOUND
    PlatformNotSupportedException = COR_E_PLATFORMNOTSUPPORTED
    SecurityException = COR_E_SECURITY
    NullReferenceException = COR_E_NULLREFERENCE
End Enum


' @References
'
' https://learn.microsoft.com/en-us/dotnet/api/system.notsupportedexception?view=netframework-4.8.1#remarks
' NotSupportedException indicates that no implementation exists for an invoked method or property.
' NotSupportedException uses the HRESULT COR_E_NOTSUPPORTED, which has the value 0x80131515.

' https://learn.microsoft.com/en-us/dotnet/api/system.typeloadexception?view=netframework-4.8.1
' TypeLoadException uses the HRESULT COR_E_TYPELOAD, that has the value 0x80131522.

' https://learn.microsoft.com/en-us/dotnet/api/system.io.fileloadexception?view=netframework-4.8.1
' FileLoadException has the default HRESULT of COR_E_FILELOAD, which has the value 0x80131621, but this is not the only possible HRESULT.

' https://learn.microsoft.com/en-us/dotnet/api/system.reflection.targetinvocationexception?view=netframework-4.8.1
' TargetInvocationException uses the HRESULT COR_E_TARGETINVOCATION which has the value 0x80131604.

' https://learn.microsoft.com/en-us/dotnet/api/system.badimageformatexception?view=netframework-4.8.1
' BadImageFormatException uses the HRESULT COR_E_BADIMAGEFORMAT, which has the value 0x8007000B.

' https://learn.microsoft.com/en-us/dotnet/api/system.indexoutofrangeexception?view=netframework-4.8.1
' IndexOutOfRangeException uses the HRESULT COR_E_INDEXOUTOFRANGE, which has the value 0x80131508.

' https://learn.microsoft.com/en-us/dotnet/api/system.invalidcastexception?view=netframework-4.8.1
' InvalidCastException uses the HRESULT COR_E_INVALIDCAST, which has the value 0x80004002.

' https://learn.microsoft.com/en-us/dotnet/api/system.io.ioexception?view=netframework-4.8.1
' IOException uses the HRESULT COR_E_IO which has the value 0x80131620.

' https://learn.microsoft.com/en-us/dotnet/api/system.unauthorizedaccessexception?view=netframework-4.8.1
' UnauthorizedAccessException uses the HRESULT COR_E_UNAUTHORIZEDACCESS, which has the value 0x80070005.

' https://learn.microsoft.com/en-us/dotnet/api/system.io.pathtoolongexception?view=netframework-4.8.1
' PathTooLongException uses the HRESULT COR_E_PATHTOOLONG, which has the value 0x800700CE

' https://learn.microsoft.com/en-us/dotnet/api/system.io.directorynotfoundexception?view=netframework-4.8.1
' DirectoryNotFoundException uses the HRESULT COR_E_DIRECTORYNOTFOUND which has the value 0x80070003

' https://learn.microsoft.com/en-us/dotnet/api/system.platformnotsupportedexception?view=netframework-4.8.1
' PlatformNotSupportedException uses the HRESULT COR_E_PLATFORMNOTSUPPORTED, which has the value 0x80131539.

' https://learn.microsoft.com/en-us/dotnet/api/system.security.securityexception?view=netframework-4.8.1
' SecurityException uses the HRESULT COR_E_SECURITY, which has the value 0x8013150A.

' https://learn.microsoft.com/en-us/dotnet/api/system.nullreferenceexception?view=netframework-4.8.1
' NullReferenceException uses the HRESULT COR_E_NULLREFERENCE, which has the value 0x80004003.
