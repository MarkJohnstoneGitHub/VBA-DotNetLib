Attribute VB_Name = "StringReplaceExample"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.replace?view=netframework-4.8.1#system-string-replace(system-string-system-string)

Option Explicit

''
' The following example demonstrates how you can use the Replace method to
' correct a spelling error.
''
Public Sub StringReplaceExample()
    Dim errString As DotNetLib.String
    Set errString = Strings.Create("This docment uses 3 other docments to docment the docmentation")
    
    Debug.Print VBString.Format("The original string is:{0}'{1}'{0}", Environment.NewLine, errString)
    
    ' Correct the spelling of "document".
    Dim correctString As DotNetLib.String
    Set correctString = errString.Replace2("docment", "document")
    
    Debug.Print VBString.Format("After correcting the string, the result is:{0}'{1}'", _
        Environment.NewLine, correctString)
End Sub

' This code example produces the following output:
'
' The original string is:
' 'This docment uses 3 other docments to docment the docmentation'
'
' After correcting the string, the result is:
' 'This document uses 3 other documents to document the documentation'
'


