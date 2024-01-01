Attribute VB_Name = "StringIsNullOrEmptyExample"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 1, 2024
'@LastModified January 1, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.isnullorempty?view=netframework-4.8.1#examples

Option Explicit

''
' The following example examines three strings and determines whether each
' string has a value, is an empty string, or is null.
''
Public Sub StringIsNullOrEmptyExample()
    Dim s1 As DotNetLib.String
    Set s1 = Strings.Create("abcd")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("")
    Dim s3 As DotNetLib.String
    Set s3 = Nothing
    
    Debug.Print VBAString.Format("String s1 {0}.", Test(s1))
    Debug.Print VBAString.Format("String s2 {0}.", Test(s2))
    Debug.Print VBAString.Format("String s3 {0}.", Test(s3))
End Sub

Private Function Test(ByVal s As DotNetLib.String) As String
    If (Strings.IsNullOrEmpty(s)) Then
        Test = "is null or empty"
    Else
        Test = VBAString.Format("(""{0}"") is neither null nor empty", s)
    End If
End Function

' The example displays the following output:
'       String s1 ("abcd") is neither null nor empty.
'       String s2 is null or empty.
'       String s3 is null or empty.
