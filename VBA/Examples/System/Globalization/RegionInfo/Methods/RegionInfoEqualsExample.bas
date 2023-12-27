Attribute VB_Name = "RegionInfoEqualsExample"
'@Folder("Examples.System.Globalization.RegionInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo.equals?view=netframework-4.8.1#examples

Option Explicit

''
' The following code example compares two instances of RegionInfo that were
' created differently.
''
Public Sub RegionInfoEquals()
    ' Creates a RegionInfo using the ISO 3166 two-letter code.
    Dim myRI1 As DotNetLib.RegionInfo
    Set myRI1 = RegionInfo.Create2("US")
    
    ' Creates a RegionInfo using a CultureInfo.LCID.
    Dim myRI2 As DotNetLib.RegionInfo
    Set myRI2 = RegionInfo.Create(CultureInfo.CreateFromName("en-US", False).LCID)

    ' Compares the two instances.
    If (myRI1.Equals(myRI2)) Then
        Debug.Print "The two RegionInfo instances are equal."
    Else
        Debug.Print "The two RegionInfo instances are NOT equal."
    End If
End Sub

'/*
'This code produces the following output.
'
'The two RegionInfo instances are equal.
'
'*/

