Attribute VB_Name = "StringSplitExample3"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 5, 2024
'@LastModified January 5, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.split?view=netframework-4.8.1#system-string-split(system-char()-system-int32)

Option Explicit

''
' The following example demonstrates how count can be used to limit the number
' of strings returned by Split.
''
Public Sub StringSplitExample3()
    Dim pvtName As DotNetLib.String
    Set pvtName = Strings.Create("Alex Johnson III")
    
    Dim subs() As String
    subs = pvtName.Split2(VBA.vbNullString, 2)
    
    Dim firstName As String
    firstName = subs(0)
    Dim lastName As String
    If UBound(subs) > 0 Then
        lastName = subs(1)
    End If
    Debug.Print VBAString.Format("firstName = {0}", firstName)
    Debug.Print VBAString.Format("lastName = {0}", lastName)
End Sub

' Output:
' firstName = Alex
' lastName = Johnson III
