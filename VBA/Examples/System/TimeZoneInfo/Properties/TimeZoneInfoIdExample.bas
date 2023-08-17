Attribute VB_Name = "TimeZoneInfoIdExample"
'@Folder "Examples.System.TimeZoneInfo.Properties"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 22, 2023
'@LastModified August 4, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo.id?view=netframework-4.8.1#examples

'@TODO Require wrapping a custom ReadOnlyCollection for SystemTime?

Option Explicit

'@Description("The following example lists the identifier of each of the time zones defined on the local computer.")
Public Sub TimeZoneInfoId()
Attribute TimeZoneInfoId.VB_Description = "The following example lists the identifier of each of the time zones defined on the local computer."
    Dim zones As DotNetLib.IReadOnlyCollection
    Set zones = TimeZoneInfo.GetSystemTimeZones()
    
    Dim varZone As Variant
    For Each varZone In zones
        Dim myZone As ITimeZoneInfo
        Set myZone = varZone
        Debug.Print myZone.Id
    Next
End Sub

' Output:
'    Dateline Standard Time
'    Utc -11
'    Aleutian Standard Time
'    Hawaiian Standard Time
'    Marquesas Standard Time
'    Alaskan Standard Time
'    Utc -9
'    Pacific Standard Time (Mexico)
'    Utc -8
'    Pacific Standard Time
'    US Mountain Standard Time
'    Mountain Standard Time (Mexico)
'    Mountain Standard Time
'    Yukon Standard Time
'    Central America Standard Time
'    Central Standard Time
'    Easter Island Standard Time
'    Central Standard Time (Mexico)
'    Canada Central Standard Time
'    SA Pacific Standard Time
'    Eastern Standard Time (Mexico)
'    Eastern Standard Time
'    Haiti Standard Time
'    Cuba Standard Time
'    US Eastern Standard Time
'    Turks And Caicos Standard Time
'    Paraguay Standard Time
'    Atlantic Standard Time
'    Venezuela Standard Time
'    Central Brazilian Standard Time
'    SA Western Standard Time
'    Pacific SA Standard Time
'    Newfoundland Standard Time
'    Tocantins Standard Time
'    E. South America Standard Time
'    SA Eastern Standard Time
'    Argentina Standard Time
'    Greenland Standard Time
'    Montevideo Standard Time
'    Magallanes Standard Time
'    Saint Pierre Standard Time
'    Bahia Standard Time
'    Utc -2
'    Mid-Atlantic Standard Time
'    Azores Standard Time
'    Cape Verde Standard Time
'    Utc
'    GMT Standard Time
'    Greenwich Standard Time
'    Sao Tome Standard Time
'    Morocco Standard Time
'    W. Europe Standard Time
'    Central Europe Standard Time
'    Romance Standard Time
'    Central European Standard Time
'    W. Central Africa Standard Time
'    GTB Standard Time
'    Middle East Standard Time
'    Egypt Standard Time
'    E. Europe Standard Time
'    Syria Standard Time
'    West Bank Standard Time
'    South Africa Standard Time
'    FLE Standard Time
'    Israel Standard Time
'    South Sudan Standard Time
'    Kaliningrad Standard Time
'    Sudan Standard Time
'    Libya Standard Time
'    Namibia Standard Time
'    Jordan Standard Time
'    Arabic Standard Time
'    Turkey Standard Time
'    Arab Standard Time
'    Belarus Standard Time
'    Russian Standard Time
'    E. Africa Standard Time
'    Volgograd Standard Time
'    Iran Standard Time
'    Arabian Standard Time
'    Astrakhan Standard Time
'    Azerbaijan Standard Time
'    Russia Time Zone 3
'    Mauritius Standard Time
'    Saratov Standard Time
'    Georgian Standard Time
'    Caucasus Standard Time
'    Afghanistan Standard Time
'    West Asia Standard Time
'    Ekaterinburg Standard Time
'    Pakistan Standard Time
'    Qyzylorda Standard Time
'    India Standard Time
'    Sri Lanka Standard Time
'    Nepal Standard Time
'    Central Asia Standard Time
'    Bangladesh Standard Time
'    Omsk Standard Time
'    Myanmar Standard Time
'    SE Asia Standard Time
'    Altai Standard Time
'    W. Mongolia Standard Time
'    North Asia Standard Time
'    N. Central Asia Standard Time
'    Tomsk Standard Time
'    China Standard Time
'    North Asia East Standard Time
'    Singapore Standard Time
'    W. Australia Standard Time
'    Taipei Standard Time
'    Ulaanbaatar Standard Time
'    Aus Central W. Standard Time
'    Transbaikal Standard Time
'    Tokyo Standard Time
'    North Korea Standard Time
'    Korea Standard Time
'    Yakutsk Standard Time
'    Cen. Australia Standard Time
'    AUS Central Standard Time
'    E. Australia Standard Time
'    AUS Eastern Standard Time
'    West Pacific Standard Time
'    Tasmania Standard Time
'    Vladivostok Standard Time
'    Lord Howe Standard Time
'    Bougainville Standard Time
'    Russia Time Zone 10
'    Magadan Standard Time
'    Norfolk Standard Time
'    Sakhalin Standard Time
'    Central Pacific Standard Time
'    Russia Time Zone 11
'    New Zealand Standard Time
'    Utc 12
'    Fiji Standard Time
'    Kamchatka Standard Time
'    Chatham Islands Standard Time
'    Utc 13
'    Tonga Standard Time
'    Samoa Standard Time
'    Line Islands Standard Time
