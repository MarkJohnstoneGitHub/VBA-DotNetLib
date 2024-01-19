Attribute VB_Name = "StringGetHashCodeExample"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 30, 2023
'@LastModified December 30, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.gethashcode?view=netframework-4.8.1#system-string-gethashcode

Option Explicit

''
' The following example demonstrates the GetHashCode method using various input strings.
''
Public Sub StringGetHashCodeExample()
    Call DisplayHashCode(Strings.Create(""))
    Call DisplayHashCode(Strings.Create("a"))
    Call DisplayHashCode(Strings.Create("ab"))
    Call DisplayHashCode(Strings.Create("abc"))
    Call DisplayHashCode(Strings.Create("abd"))
    Call DisplayHashCode(Strings.Create("abe"))
    Call DisplayHashCode(Strings.Create("abcdef"))
    Call DisplayHashCode(Strings.Create("abcdeg"))
    Call DisplayHashCode(Strings.Create("abcdeh"))
    Call DisplayHashCode(Strings.Create("abcdei"))
    Call DisplayHashCode(Strings.Create("Abcdeg"))
    Call DisplayHashCode(Strings.Create("Abcdeh"))
    Call DisplayHashCode(Strings.Create("Abcdei"))
End Sub

Private Sub DisplayHashCode(ByVal Operand As DotNetLib.String)
    Dim pvtHashCode As Long
     pvtHashCode = Operand.GetHashCode()
    Debug.Print VBString.Format("The hash code for ""{0}"" is: 0x{1:X8}, {1}", _
                         Operand, pvtHashCode)
End Sub

'/*
'      This example displays output like the following:
'      The hash code for "" is: 0x2D2816FE, 757602046
'      The hash code for "a" is: 0xCDCAB7BF, -842352705
'      The hash code for "ab" is: 0xCDE8B7BF, -840386625
'      The hash code for "abc" is: 0x2001D81A, 536991770
'      The hash code for "abd" is: 0xC2A94CB5, -1029092171
'      The hash code for "abe" is: 0x6550C150, 1699791184
'      The hash code for "abcdef" is: 0x1762906D, 392335469
'      The hash code for "abcdeg" is: 0x1763906D, 392401005
'      The hash code for "abcdeh" is: 0x175C906D, 391942253
'      The hash code for "abcdei" is: 0x175D906D, 392007789
'      The hash code for "Abcdeg" is: 0x1763954D, 392402253
'      The hash code for "Abcdeh" is: 0x175C954D, 391943501
'      The hash code for "Abcdei" is: 0x175D954D, 392009037
'*/


