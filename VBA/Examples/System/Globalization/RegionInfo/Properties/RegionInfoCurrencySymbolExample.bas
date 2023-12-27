Attribute VB_Name = "RegionInfoCurrencySymbolExample"
'@Folder("Examples.System.Globalization.RegionInfo.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo.currencysymbol?view=netframework-4.8.1#examples

Option Explicit

Public Sub RegionInfoCurrencySymbolExample()
    ' Displays the property values of the RegionInfo for "US".
    Dim myRI1 As DotNetLib.RegionInfo
    Set myRI1 = RegionInfo.Create2("US")
    Debug.Print VBAString.Format("   Name:                         {0}", myRI1.Name)
    Debug.Print VBAString.Format("   DisplayName:                  {0}", myRI1.DisplayName)
    Debug.Print VBAString.Format("   EnglishName:                  {0}", myRI1.EnglishName)
    Debug.Print VBAString.Format("   IsMetric:                     {0}", myRI1.IsMetric)
    Debug.Print VBAString.Format("   ThreeLetterISORegionName:     {0}", myRI1.ThreeLetterISORegionName)
    Debug.Print VBAString.Format("   ThreeLetterWindowsRegionName: {0}", myRI1.ThreeLetterWindowsRegionName)
    Debug.Print VBAString.Format("   TwoLetterISORegionName:       {0}", myRI1.TwoLetterISORegionName)
    Debug.Print VBAString.Format("   CurrencySymbol:               {0}", myRI1.CurrencySymbol)
    Debug.Print VBAString.Format("   ISOCurrencySymbol:            {0}", myRI1.ISOCurrencySymbol)
End Sub

'/*
'This code produces the following output.
'
'Name:                            US
'DisplayName:                     United States
'EnglishName:                     United States
'   IsMetric:                     False
'ThreeLetterISORegionName:        USA
'ThreeLetterWindowsRegionName:    USA
'TwoLetterISORegionName:          US
'   CurrencySymbol:               $
'ISOCurrencySymbol:               USD
'
'*/
