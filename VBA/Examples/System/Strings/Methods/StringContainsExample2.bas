Attribute VB_Name = "StringContainsExample2"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.contains?view=netframework-4.8.1#system-string-contains(system-string-system-stringcomparison)

Option Explicit

Public Sub StringContainsExample2()
    Dim s As DotNetLib.String
    Set s = Strings.Create("This is a string.")
    Dim sub1 As DotNetLib.String
    Set sub1 = Strings.Create("this")
    
    Debug.Print VBString.Format("Does '{0}' contain '{1}'?", s, sub1)
    
    Dim comp As mscorlib.StringComparison
    comp = StringComparison.StringComparison_Ordinal
    Debug.Print VBString.Format("   {0:G}: {1}", StringComparisonHelper.ToString(comp), s.Contains(sub1, comp))
    
    comp = StringComparison.StringComparison_OrdinalIgnoreCase
    Debug.Print VBString.Format("   {0:G}: {1}", StringComparisonHelper.ToString(comp), s.Contains(sub1, comp))
End Sub

' The example displays the following output:
'       Does 'This is a string.' contain 'this'?
'          Ordinal: False
'          OrdinalIgnoreCase: True
