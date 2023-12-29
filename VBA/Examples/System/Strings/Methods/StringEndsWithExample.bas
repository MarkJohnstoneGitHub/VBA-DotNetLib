Attribute VB_Name = "StringEndsWithExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.endswith?view=netframework-4.8.1#system-string-endswith(system-string)

Option Explicit

Public Sub StringEndsWithExample()
    Dim pvtStrings() As String
    Call ArrayEx.CreateInitialize1D(pvtStrings, "This is a string.", "Hello!", "Nothing.", _
                                    "Yes.", "randomize")
    Dim value As Variant
    For Each value In pvtStrings
        Dim endsInPeriod As Boolean
        Dim pvtString As DotNetLib.String
        Set pvtString = Strings.Create(value)
        endsInPeriod = pvtString.EndsWith3(".")
        Debug.Print VBAString.Format("'{0}' ends in a period: {1}", _
                              pvtString, endsInPeriod)
    Next
End Sub

' The example displays the following output:
'       'This is a string.' ends in a period: True
'       'Hello!' ends in a period: False
'       'Nothing.' ends in a period: True
'       'Yes.' ends in a period: True
'       'randomize' ends in a period: False

