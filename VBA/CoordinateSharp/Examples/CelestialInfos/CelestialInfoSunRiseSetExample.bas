Attribute VB_Name = "CelestialInfoSunRiseSetExample"
'@Folder("CoordinateSharp.Examples.CelestialInfos")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 19, 2023
'@LastModified August 19, 2023

' https://www.nuget.org/packages/CoordinateSharp/#readme-body-tab
' https://coordinatesharp.com/Help/html/T_CoordinateSharp_Celestial.htm
' https://stackoverflow.com/questions/76929620/calculating-sunrise-sunset-in-vba#76929620

Option Explicit

Private Type CoOrdinate
    latitude As Double
    longitude As Double
End Type

Public Sub CelestialInfoExample()
    'Melbourne Australia
    Dim melbourneTimeZoneId As String
    melbourneTimeZoneId = "AUS Eastern Standard Time"
    
    Dim melboureCoOrdinate As CoOrdinate
    melboureCoOrdinate.latitude = -37.840935
    melboureCoOrdinate.longitude = 144.946457
    
    Dim utcSunRise As DotNetLib.DateTime
    Dim localSunRise As DotNetLib.DateTime
    Set utcSunRise = CelestialInfo.Sunrise(melboureCoOrdinate.latitude, melboureCoOrdinate.longitude, DateTime.Today)
    Set localSunRise = TimeZoneInfo.ConvertTimeFromUtc(utcSunRise, TimeZoneInfo.FindSystemTimeZoneById(melbourneTimeZoneId))
    
    Dim utcSunSet As DotNetLib.DateTime
    Dim localSunSet As DotNetLib.DateTime
    
    Set utcSunSet = CelestialInfo.SunSet(melboureCoOrdinate.latitude, melboureCoOrdinate.longitude, DateTime.Today)
    Set localSunSet = TimeZoneInfo.ConvertTimeFromUtc(utcSunSet, TimeZoneInfo.FindSystemTimeZoneById(melbourneTimeZoneId))
    
    Debug.Print "Sunrise in Melbourne : "; localSunRise.ToString()
    Debug.Print "Sunset in Melbourne  : "; localSunSet.ToString()
End Sub

'Output:
'    Sunrise in Melbourne : 20/08/2023 7:01:13 AM
'    Sunset in Melbourne  : 19/08/2023 5:48:28 PM
