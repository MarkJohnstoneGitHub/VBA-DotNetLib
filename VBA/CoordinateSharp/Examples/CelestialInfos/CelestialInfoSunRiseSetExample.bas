Attribute VB_Name = "CelestialInfoSunRiseSetExample"
Attribute VB_Description = "Celestial Info example displaying sunrise and sunset for local time."
'@ModuleDescription "Celestial Info example displaying sunrise and sunset for local time."
'@Folder("CoordinateSharp.Examples.CelestialInfos")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 19, 2023
'@LastModified September 3, 2023

' https://www.nuget.org/packages/CoordinateSharp/#readme-body-tab
' https://coordinatesharp.com/Help/html/T_CoordinateSharp_Celestial.htm
' https://stackoverflow.com/questions/76929620/calculating-sunrise-sunset-in-vba#76929620
' https://www.timeanddate.com/sun/australia/melbourne

Option Explicit

Private Type CoOrdinate
    latitude As Double
    longitude As Double
End Type

'@Description("Melbourne, Australia sunrise and sunset for local time.")
'@Reference https://www.timeanddate.com/sun/australia/melbourne
'@Bug Displaying sunrise for following day? Possible CoordinateSharp bug to report.
Public Sub CelestialInfoExample()
Attribute CelestialInfoExample.VB_Description = "Melbourne, Australia sunrise and sunset for local time."
    Dim melbourneTimeZoneId As String
    melbourneTimeZoneId = "AUS Eastern Standard Time"
    
    Dim melboureCoOrdinate As CoOrdinate
    melboureCoOrdinate.latitude = -37.840935
    melboureCoOrdinate.longitude = 144.946457
    
    Dim melbourneTZI As DotNetLib.TimeZoneInfo
    Set melbourneTZI = TimeZoneInfo.FindSystemTimeZoneById(melbourneTimeZoneId)
    
    Dim dto As DotNetLib.DateTimeOffset
    Set dto = DateTimeOffset.CreateFromDateTime(DateTime.UtcNow)
    
    Dim utcSunRise As DotNetLib.DateTime
    Set utcSunRise = CelestialInfo.SunRise(melboureCoOrdinate.latitude, melboureCoOrdinate.longitude, dto.UtcDateTime)
    Dim localSunRise As DotNetLib.DateTime
    Set localSunRise = TimeZoneInfo.ConvertTimeFromUtc(utcSunRise, melbourneTZI)
    Debug.Print "Melbourne,Australia SunRise : "; localSunRise.ToString
    
    Dim utcSunSet As DotNetLib.DateTime
    Set utcSunSet = CelestialInfo.SunSet(melboureCoOrdinate.latitude, melboureCoOrdinate.longitude, dto.UtcDateTime)
    Dim localSunSet As DotNetLib.DateTime
    Set localSunSet = TimeZoneInfo.ConvertTimeFromUtc(utcSunSet, melbourneTZI)
    Debug.Print "Melbourne,Australia SunSet :  "; localSunSet.ToString
    
    Dim daylength As DotNetLib.TimeSpan
    Dim nightLength As DotNetLib.TimeSpan
    
    Set nightLength = DateTime.Subtraction(utcSunRise, utcSunSet)
    Set daylength = TimeSpan.Subtraction(TimeSpan.Create(24, 0, 0), nightLength)
    
    Debug.Print "Day lenth "; daylength.ToString()
    Debug.Print "Night length: "; nightLength.ToString()
End Sub

' Output for August 25, 2023
'    Melbourne,Australia SunRise : 26/08/2023 6:53:09 AM
'    Melbourne,Australia SunSet :  25/08/2023 5:53:34 PM
'    Day lenth 11:00:25
'    Night length: 12:59:35


'@Description("New York City, USA sunrise and sunset for local time.")
'@Reference https://www.timeanddate.com/sun/usa/new-york
Public Sub CelestialInfoExample2()
Attribute CelestialInfoExample2.VB_Description = "New York City, USA sunrise and sunset for local time."
    Dim nycTimeZoneId As String
    nycTimeZoneId = "Eastern Standard Time"
    
    Dim nycCoOrdinate As CoOrdinate
    nycCoOrdinate.latitude = 40.73061
    nycCoOrdinate.longitude = -73.935242
    Dim nycTZI As DotNetLib.TimeZoneInfo
    Set nycTZI = TimeZoneInfo.FindSystemTimeZoneById(nycTimeZoneId)
    
    Dim dto As DotNetLib.DateTimeOffset
    Set dto = DateTimeOffset.CreateFromDateTime(DateTime.UtcNow)
    
    Dim utcSunRise As DotNetLib.DateTime
    Set utcSunRise = CelestialInfo.SunRise(nycCoOrdinate.latitude, nycCoOrdinate.longitude, dto.UtcDateTime)
    Dim localSunRise As DotNetLib.DateTime
    Set localSunRise = TimeZoneInfo.ConvertTimeFromUtc(utcSunRise, nycTZI)
    Debug.Print "New York City sunrise : "; localSunRise.ToString
    
    Dim utcSunSet As DotNetLib.DateTime
    Set utcSunSet = CelestialInfo.SunSet(nycCoOrdinate.latitude, nycCoOrdinate.longitude, dto.UtcDateTime)
    Dim localSunSet As DotNetLib.DateTime
    Set localSunSet = TimeZoneInfo.ConvertTimeFromUtc(utcSunSet, nycTZI)
    Debug.Print "New York City sunset :  "; localSunSet.ToString
    
    Dim daylength As DotNetLib.TimeSpan
    Dim nightLength As DotNetLib.TimeSpan
    
    Set daylength = DateTime.Subtraction(utcSunSet, utcSunRise)
    Set nightLength = TimeSpan.Subtraction(TimeSpan.Create(24, 0, 0), daylength)
    
    Debug.Print "Day lenth "; daylength.ToString()
    Debug.Print "Night length: "; nightLength.ToString()
End Sub

' Output for August 25, 2023
'    New York City sunrise : 25/08/2023 6:16:59 AM
'    New York City sunset :  25/08/2023 7:41:51 PM
'    Day lenth 13:24:52
'    Night length: 10:35:08


'@Description("Amsterdam, Netherlands sunrise and sunset for local time.")
'@Reference https://www.timeanddate.com/sun/netherlands/amsterdam
Public Sub CelestialInfoExample3()
Attribute CelestialInfoExample3.VB_Description = "Amsterdam, Netherlands sunrise and sunset for local time."
    Const AmsterdamTimeZoneId As String = "W. Europe Standard Time"
    
    Dim amsterdamCoOrdinate As CoOrdinate
    amsterdamCoOrdinate.latitude = 52.3676
    amsterdamCoOrdinate.longitude = 4.9041
    
    Dim amsterdamTZI As DotNetLib.TimeZoneInfo
    Set amsterdamTZI = TimeZoneInfo.FindSystemTimeZoneById(AmsterdamTimeZoneId)
    Dim dto As DotNetLib.DateTimeOffset
    Set dto = DateTimeOffset.CreateFromDateTime(DateTime.UtcNow)
    
    Dim utcSunRise As DotNetLib.DateTime
    Set utcSunRise = CelestialInfo.SunRise(amsterdamCoOrdinate.latitude, amsterdamCoOrdinate.longitude, dto.UtcDateTime)
    Dim localSunRise As DotNetLib.DateTime
    Set localSunRise = TimeZoneInfo.ConvertTimeFromUtc(utcSunRise, amsterdamTZI)
    Debug.Print "Amsterdam, Netherlands sunrise : "; localSunRise.ToString
    
    Dim utcSunSet As DotNetLib.DateTime
    Set utcSunSet = CelestialInfo.SunSet(amsterdamCoOrdinate.latitude, amsterdamCoOrdinate.longitude, dto.UtcDateTime)
    Dim localSunSet As DotNetLib.DateTime
    Set localSunSet = TimeZoneInfo.ConvertTimeFromUtc(utcSunSet, amsterdamTZI)
    Debug.Print "Amsterdam, Netherlands sunset  : "; localSunSet.ToString
    
    Dim daylength As DotNetLib.TimeSpan
    Dim nightLength As DotNetLib.TimeSpan
    
    Set daylength = DateTime.Subtraction(utcSunSet, utcSunRise)
    Set nightLength = TimeSpan.Subtraction(TimeSpan.Create(24, 0, 0), daylength)
    
    Debug.Print "Day lenth "; daylength.ToString()
    Debug.Print "Night length: "; nightLength.ToString()
End Sub

' Output for August 25, 2023
'    Amsterdam, Netherlands sunrise : 25/08/2023 6:40:32 AM
'    Amsterdam, Netherlands sunset  : 25/08/2023 8:47:42 PM
'    Day lenth 14:07:10
'    Night length: 09:52:50
