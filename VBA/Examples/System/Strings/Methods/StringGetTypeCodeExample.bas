Attribute VB_Name = "StringGetTypeCodeExample"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.gettypecode?view=netframework-4.8.1

Option Explicit

''
' The following example displays the TypeCode enumerated constant for the String type.
''
Public Sub StringGetTypeCodeExample()
    Dim str As DotNetLib.String
    Set str = Strings.Create("abc")
    Dim tc As mscorlib.TypeCode
    tc = str.GetTypeCode()
    Debug.Print VBString.Format("The type code for '{0}' is {1}, which represents {2}.", _
                         str, TypeCodeHelper.ToString2(tc, "D"), TypeCodeHelper.ToString2(tc, "F"))
End Sub

'/*
'This example produces the following results:
'The type code for 'abc' is 18, which represents String.
'*/


