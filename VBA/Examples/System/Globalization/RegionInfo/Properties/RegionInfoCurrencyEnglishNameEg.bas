Attribute VB_Name = "RegionInfoCurrencyEnglishNameEg"
'@Folder "Examples.System.Globalization.RegionInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo.currencyenglishname?view=netframework-4.8.1#examples

Option Explicit

''
' This example demonstrates the RegionInfo.EnglishName, NativeName,
' CurrencyEnglishName, CurrencyNativeName, and GeoId properties.
''
Public Sub RegionInfoCurrencyEnglishNameExample()
    Dim ri As DotNetLib.RegionInfo
    Set ri = RegionInfo.Create2("SE") ' Sweden

    Debug.Print VBString.Format("Region English Name: . . . {0}", ri.EnglishName)
    Debug.Print VBString.Format("Native Name: . . . . . . . {0}", ri.NativeName)
    Debug.Print VBString.Format("Currency English Name: . . {0}", ri.CurrencyEnglishName)
    Debug.Print VBString.Format("Currency Native Name:. . . {0}", ri.CurrencyNativeName)
    Debug.Print VBString.Format("Geographical ID: . . . . . {0}", ri.GeoId)
End Sub

'This code example produces the following results:
'
'Region English Name: . . . Sweden
'Native Name: . . . . . . . Sverige
'Currency English Name: . . Swedish Krona
'Currency Native Name:. . . Svensk krona
'Geographical ID: . . . . . 221
'
'*/
