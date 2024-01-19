Attribute VB_Name = "VBAStringEndsWithExample"
'@Folder "Examples.System.VBString.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified January 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.endswith?view=netframework-4.8.1#system-string-endswith(system-string)

Option Explicit

Public Sub VBAStringEndsWithExample()
    Dim pvtStrings() As String
    Call VBArray.CreateInitialize1D(pvtStrings, "This is a string.", "Hello!", "Nothing.", _
                                    "Yes.", "randomize")
    Dim value As Variant
    For Each value In pvtStrings
        Dim endsInPeriod As Boolean
        endsInPeriod = VBString.EndsWith(CStr(value), ".")
        Debug.Print VBString.Format("'{0}' ends in a period: {1}", _
                              value, endsInPeriod)
    Next
End Sub

' The example displays the following output:
'       'This is a string.' ends in a period: True
'       'Hello!' ends in a period: False
'       'Nothing.' ends in a period: True
'       'Yes.' ends in a period: True
'       'randomize' ends in a period: False


