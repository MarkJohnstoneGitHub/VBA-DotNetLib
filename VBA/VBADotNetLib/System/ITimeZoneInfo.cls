VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITimeZoneInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "VBADotNetLib.System"
'@Interface

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 30, 2023
'@LastModified July 31, 2023

Option Explicit

'@Description("Gets the time difference between the current time zone's standard time and Coordinated Universal Time (UTC).")
Public Property Get BaseUtcOffset() As TimeSpan
Attribute BaseUtcOffset.VB_Description = "Gets the time difference between the current time zone's standard time and Coordinated Universal Time (UTC)."
End Property

'@Description("Gets the display name for the current time zone's daylight saving time.")
Public Property Get DaylightName() As String
Attribute DaylightName.VB_Description = "Gets the display name for the current time zone's daylight saving time."
End Property

'@Description("Gets the time zone identifier.")
Public Property Get DisplayName() As String
Attribute DisplayName.VB_Description = "Gets the time zone identifier."
End Property

'@Description("Gets the time zone identifier.")
Public Property Get Id() As String
Attribute Id.VB_Description = "Gets the time zone identifier."
End Property

'@Description("Gets the display name for the time zone's standard time.")
Public Property Get StandardName() As String
Attribute StandardName.VB_Description = "Gets the display name for the time zone's standard time."
End Property

'@Description("Gets the display name for the time zone's standard time.")
Public Property Get SupportsDaylightSavingTime() As Boolean
Attribute SupportsDaylightSavingTime.VB_Description = "Gets the display name for the time zone's standard time."
End Property

'@Description("Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal.")
Public Function Equals(ByVal other As TimeZoneInfo) As Boolean
Attribute Equals.VB_Description = "Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal."
End Function

'@Description("Determines whether the current TimeZoneInfo object and another object are equal.")
Public Function Equals2(ByVal obj As Variant) As Boolean
Attribute Equals2.VB_Description = "Determines whether the current TimeZoneInfo object and another object are equal."
End Function

'@Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")
Public Function GetAmbiguousTimeOffsets(ByVal sourceDateTime As DateTime) As TimeSpan()
Attribute GetAmbiguousTimeOffsets.VB_Description = "Returns information about the possible dates and times that an ambiguous date and time can be mapped to."
End Function

'@Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")
Public Function GetAmbiguousTimeOffsets2(ByVal sourceDateTimeOffset As DateTimeOffset) As TimeSpan()
Attribute GetAmbiguousTimeOffsets2.VB_Description = "Returns information about the possible dates and times that an ambiguous date and time can be mapped to."
End Function

'@Description("Serves as a hash function for hashing algorithms and data structures such as hash tables.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Serves as a hash function for hashing algorithms and data structures such as hash tables."
End Function

'@Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")
Public Function GetUtcOffset(ByVal sourceDateTime As DateTime) As TimeSpan
Attribute GetUtcOffset.VB_Description = "Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time."
End Function

'@Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")
Public Function GetUtcOffset2(ByVal sourceDateTimeOffset As DateTimeOffset) As TimeSpan
Attribute GetUtcOffset2.VB_Description = "Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time."
End Function

'@Description("Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules.")
Public Function HasSameRules(ByVal other As TimeZoneInfo) As Boolean
Attribute HasSameRules.VB_Description = "Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules."
End Function

'@Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")
Public Function IsAmbiguousTime(ByVal sourceDateTime As DateTime) As Boolean
Attribute IsAmbiguousTime.VB_Description = "Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times."
End Function

'@Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")
Public Function IsAmbiguousTime2(ByVal sourceDateTimeOffset As DateTimeOffset) As Boolean
Attribute IsAmbiguousTime2.VB_Description = "Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times."
End Function

'@Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")
Public Function IsDaylightSavingTime(ByVal sourceDateTime As DateTime) As Boolean
Attribute IsDaylightSavingTime.VB_Description = "Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object."
End Function

'@Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")
Public Function IsDaylightSavingTime2(ByVal sourceDateTimeOffset As DateTimeOffset) As Boolean
Attribute IsDaylightSavingTime2.VB_Description = "Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object."
End Function

'@Description("Indicates whether a particular date and time is invalid.")
Public Function IsInvalidTime(ByVal sourceDateTime As DateTime) As Boolean
Attribute IsInvalidTime.VB_Description = "Indicates whether a particular date and time is invalid."
End Function

'@Description("Converts the current TimeZoneInfo object to a serialized string.")
Public Function ToSerializedString() As String
Attribute ToSerializedString.VB_Description = "Converts the current TimeZoneInfo object to a serialized string."
End Function

'@Description("Returns the current TimeZoneInfo object's display name.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Returns the current TimeZoneInfo object's display name."
End Property
