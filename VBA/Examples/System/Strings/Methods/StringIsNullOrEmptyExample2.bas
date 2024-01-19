Attribute VB_Name = "StringIsNullOrEmptyExample2"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 1, 2024
'@LastModified January 1, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.isnullorempty?view=netframework-4.8.1#remarks

Option Explicit

''
' IsNullOrEmpty is a convenience method that enables you to simultaneously test
' whether a String is null or its value is String.Empty.
' It is equivalent to the following code:
''
Public Sub StringIsNullOrEmptyExample2()
    Dim s1 As DotNetLib.String
    Set s1 = Nothing
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("")
    Dim s3 As DotNetLib.String
    Set s3 = Strings.Create(VBA.vbNullString)
    Dim str As String
    Dim s4 As DotNetLib.String
    Set s4 = Strings.Create(str)
    Debug.Print TestForNullOrEmpty(s1)
    Debug.Print TestForNullOrEmpty(s2)
    Debug.Print TestForNullOrEmpty(s3)
    Debug.Print TestForNullOrEmpty(s4)
End Sub

Private Function TestForNullOrEmpty(ByVal s As DotNetLib.String) As Boolean
    Dim result As Boolean
    result = (s Is Nothing) Or Strings.Equals(s, Strings.EmptyString)
    TestForNullOrEmpty = result
End Function

' The example displays the following output:
'    True
'    True
'    True
'    True
