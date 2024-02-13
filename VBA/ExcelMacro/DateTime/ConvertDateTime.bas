Attribute VB_Name = "ConvertDateTime"
'@Folder("ExcelMacro.DateTime")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 14, 2024
'@LastModified February 14, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

Option Explicit

''
' Obtains the current UTC Date time.
''
Public Function GetUTCNow() As Date
    GetUTCNow = DateTime.UtcNow.ToOADate
End Function

''
' Converts a UTC Date time to local time
''
Public Function ConvertToLocalTime(ByVal utcDate As Date) As Date
    ConvertToLocalTime = DateTime.FromOADate(utcDate).ToLocalTime.ToOADate
End Function

