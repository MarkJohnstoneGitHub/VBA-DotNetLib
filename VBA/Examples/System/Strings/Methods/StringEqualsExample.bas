Attribute VB_Name = "StringEqualsExample"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.equals?view=netframework-4.8.1#system-string-equals(system-object)

Option Explicit

''
' The following example demonstrates the String.Equals method.
''
Public Sub StringEqualsExample()
    Dim sb As DotNetLib.StringBuilder
    Set sb = StringBuilder.Create("abcd")
    Dim str1 As DotNetLib.String
    Set str1 = Strings.Create("abcd")
    Dim str2 As DotNetLib.String
    Set str2 = Strings.Create(VBA.vbNullString)
    Dim o2 As Object
    Set o2 = Nothing
    
    Debug.Print
    Debug.Print VBString.Format(" *  The value of String str1 is '{0}'.", str1)
    Debug.Print VBString.Format(" *  The value of StringBuilder sb is '{0}'.", sb.ToString())

    Debug.Print
    Debug.Print VBString.Format("1a) String.Equals(Object). Object is a StringBuilder, not a String.")
    Debug.Print VBString.Format("    Is str1 equal to sb?: {0}", str1.Equals(sb))
    
    Debug.Print
    Debug.Print VBString.Format("1b) String.Equals(Object). Object is a String.");
    Set str2 = Strings.Create(sb.ToString())
    Set o2 = str2
    Debug.Print VBString.Format(" *  The value of Object o2 is '{0}'.", o2)
    Debug.Print VBString.Format("    Is str1 equal to o2?: {0}", str1.Equals(o2))

    Debug.Print
    Debug.Print VBString.Format(" 2) String.Equals(String)")
    Debug.Print VBString.Format(" *  The value of String str2 is '{0}'.", str2)
    Debug.Print VBString.Format("    Is str1 equal to str2?: {0}", str1.Equals2(str2))

    Debug.Print
    Debug.Print VBString.Format(" 3) String.Equals(String, String)")
    Debug.Print VBString.Format("    Is str1 equal to str2?: {0}", Strings.Equals(str1, str2))
End Sub

'/*
'This example produces the following results:
'
' *  The value of String str1 is 'abcd'.
' *  The value of StringBuilder sb is 'abcd'.
'
'1a) String.Equals(Object). Object is a StringBuilder, not a String.
'    Is str1 equal to sb?: False
'
'1b) String.Equals(Object). Object is a String.
' *  The value of Object o2 is 'abcd'.
'    Is str1 equal to o2?: True
'
' 2) String.Equals(String)
' *  The value of String str2 is 'abcd'.
'    Is str1 equal to str2?: True
'
' 3) String.Equals(String, String)
'    Is str1 equal to str2?: True
'*/
