Attribute VB_Name = "StringIsNullOrEmptyExample"
'@IgnoreModule VariableNotAssigned, EmptyStringLiteral
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 1, 2024
'@LastModified January 28, 2024

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
    
    Dim s4 As DotNetLib.String
    Set s4 = Strings.Create(VBA.vbNullString)
    Dim str As String
    Dim s5 As DotNetLib.String
    Set s5 = Strings.Create(str)
    
    Debug.Print VBString.Format("String s1 {0}.", test(s1))
    Debug.Print VBString.Format("String s2 {0}.", test(s2))
    Debug.Print VBString.Format("String s3 {0}.", test(s3))
    Debug.Print VBString.Format("String s4 {0}.", test(s4))
    Debug.Print VBString.Format("String s5 {0}.", test(s5))
End Sub

Private Function test(ByVal s As DotNetLib.String) As String
    If (Strings.IsNullOrEmpty(s)) Then
        test = "is null or empty"
    Else
        test = VBString.Format("(""{0}"") is neither null nor empty", s)
    End If
End Function

' The example displays the following output:
'       String s1 ("abcd") is neither null nor empty.
'       String s2 is null or empty.
'       String s3 is null or empty.
'       String s4 is null or empty.
'       String s5 is null or empty.

