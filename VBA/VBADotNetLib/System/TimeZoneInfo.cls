VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeZoneInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents any time zone in the world."
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("Represents any time zone in the world.")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 11, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1

'@Remarks
' Generated from DotNetLib.tlb using https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class
' Renamed DotNetlib.TimeZoneInfo.Local to Locale due to VBA reserved word


Option Explicit

Private Type TTimeZoneInfo
   DotNetLibTimeZoneInfo As DotNetLib.TimeZoneInfo
End Type

Private this As TTimeZoneInfo

Private Sub Class_Initialize()
   Set this.DotNetLibTimeZoneInfo = New DotNetLib.TimeZoneInfo
End Sub

Private Sub Class_Terminate()
   Set this.DotNetLibTimeZoneInfo = Nothing
End Sub

Friend Property Get ComObject() As DotNetLib.TimeZoneInfo
   Set ComObject = this.DotNetLibTimeZoneInfo
End Property

Friend Property Set ComObject(ByVal objTimeZoneInfo As DotNetLib.TimeZoneInfo)
   Set this.DotNetLibTimeZoneInfo = objTimeZoneInfo
End Property

Friend Property Get Self() As TimeZoneInfo
   Set Self = Me
End Property

'@Description("Gets the time difference between the current time zone's standard time and Coordinated Universal Time (UTC).")
Public Property Get BaseUtcOffset() As TimeSpan
Attribute BaseUtcOffset.VB_Description = "Gets the time difference between the current time zone's standard time and Coordinated Universal Time (UTC)."
   With New TimeSpan
      Set .ComObject = this.DotNetLibTimeZoneInfo.BaseUtcOffset()
      Set BaseUtcOffset = .Self
   End With
End Property

'@Description("Gets the display name for the current time zone's daylight saving time.")
Public Property Get DaylightName() As String
Attribute DaylightName.VB_Description = "Gets the display name for the current time zone's daylight saving time."
   DaylightName = this.DotNetLibTimeZoneInfo.DaylightName()
End Property

'@Description("Gets the general display name that represents the time zone.")
Public Property Get DisplayName() As String
Attribute DisplayName.VB_Description = "Gets the general display name that represents the time zone."
   DisplayName = this.DotNetLibTimeZoneInfo.DisplayName()
End Property

'@Description("Gets the time zone identifier.")
Public Property Get Id() As String
Attribute Id.VB_Description = "Gets the time zone identifier."
   Id = this.DotNetLibTimeZoneInfo.Id()
End Property

'@Description("Gets a TimeZoneInfo object that represents the local time zone.")
Public Property Get Locale() As TimeZoneInfo
Attribute Locale.VB_Description = "Gets a TimeZoneInfo object that represents the local time zone."
   With New TimeZoneInfo
      Set .ComObject = this.DotNetLibTimeZoneInfo.Local()
      Set Locale = .Self
   End With
End Property

'@Description("Gets the display name for the time zone's standard time.")
Public Property Get StandardName() As String
Attribute StandardName.VB_Description = "Gets the display name for the time zone's standard time."
   StandardName = this.DotNetLibTimeZoneInfo.StandardName()
End Property

'@Description("Gets a value indicating whether the time zone has any daylight saving time rules.")
Public Property Get SupportsDaylightSavingTime() As Boolean
Attribute SupportsDaylightSavingTime.VB_Description = "Gets a value indicating whether the time zone has any daylight saving time rules."
   SupportsDaylightSavingTime = this.DotNetLibTimeZoneInfo.SupportsDaylightSavingTime()
End Property

'@Description("Gets a TimeZoneInfo object that represents the Coordinated Universal Time (UTC) zone.")
Public Property Get Utc() As TimeZoneInfo
Attribute Utc.VB_Description = "Gets a TimeZoneInfo object that represents the Coordinated Universal Time (UTC) zone."
   With New TimeZoneInfo
      Set .ComObject = this.DotNetLibTimeZoneInfo.Utc()
      Set Utc = .Self
   End With
End Property

'@Description("Clears cached time zone data.")
Public Sub ClearCachedData()
Attribute ClearCachedData.VB_Description = "Clears cached time zone data."
   this.DotNetLibTimeZoneInfo.ClearCachedData
End Sub

'@Static
'@Description("Converts a time to the time in a particular time zone.")
Public Function ConvertTime(ByRef DateTime As DateTime, ByRef destinationTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTime.VB_Description = "Converts a time to the time in a particular time zone."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTime(DateTime.ComObject, destinationTimeZone.ComObject)
      Set ConvertTime = .Self
   End With
End Function

'@Static
'@Description("Converts a time to the time in a particular time zone.")
Public Function ConvertTime2(ByRef DateTimeOffset As DateTimeOffset, ByRef destinationTimeZone As TimeZoneInfo) As DateTimeOffset
Attribute ConvertTime2.VB_Description = "Converts a time to the time in a particular time zone."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTime2(DateTimeOffset.ComObject, destinationTimeZone.ComObject)
      Set ConvertTime2 = .Self
   End With
End Function

'@Static
'@Description("Converts a time from one time zone to another.")
Public Function ConvertTime3(ByRef DateTime As DateTime, ByRef sourceTimeZone As TimeZoneInfo, ByRef destinationTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTime3.VB_Description = "Converts a time from one time zone to another."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTime3(DateTime.ComObject, sourceTimeZone.ComObject, destinationTimeZone.ComObject)
      Set ConvertTime3 = .Self
   End With
End Function

'@Static
'@Description("Converts a time to the time in another time zone based on the time zone's identifier.")
Public Function ConvertTimeBySystemTimeZoneId(ByRef DateTime As DateTime, ByVal destinationTimeZoneId As String) As DateTime
Attribute ConvertTimeBySystemTimeZoneId.VB_Description = "Converts a time to the time in another time zone based on the time zone's identifier."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime.ComObject, destinationTimeZoneId)
      Set ConvertTimeBySystemTimeZoneId = .Self
   End With
End Function

'@Static
'@Description("Converts a time to the time in another time zone based on the time zone's identifier.")
Public Function ConvertTimeBySystemTimeZoneId2(ByRef DateTimeOffset As DateTimeOffset, ByVal destinationTimeZoneId As String) As DateTimeOffset
Attribute ConvertTimeBySystemTimeZoneId2.VB_Description = "Converts a time to the time in another time zone based on the time zone's identifier."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeBySystemTimeZoneId2(DateTimeOffset.ComObject, destinationTimeZoneId)
      Set ConvertTimeBySystemTimeZoneId2 = .Self
   End With
End Function

'@Static
'@Description("Converts a time from one time zone to another based on time zone identifiers.")
Public Function ConvertTimeBySystemTimeZoneId3(ByRef DateTime As DateTime, ByVal sourceTimeZoneId As String, ByVal destinationTimeZoneId As String) As DateTime
Attribute ConvertTimeBySystemTimeZoneId3.VB_Description = "Converts a time from one time zone to another based on time zone identifiers."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeBySystemTimeZoneId3(DateTime.ComObject, sourceTimeZoneId, destinationTimeZoneId)
      Set ConvertTimeBySystemTimeZoneId3 = .Self
   End With
End Function

'@Static
'@Description("Converts a Coordinated Universal Time (UTC) to the time in a specified time zone.")
Public Function ConvertTimeFromUtc(ByRef DateTime As DateTime, ByRef destinationTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTimeFromUtc.VB_Description = "Converts a Coordinated Universal Time (UTC) to the time in a specified time zone."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeFromUtc(DateTime.ComObject, destinationTimeZone.ComObject)
      Set ConvertTimeFromUtc = .Self
   End With
End Function

'@Static
'@Description("Converts the specified date and time to Coordinated Universal Time (UTC).")
Public Function ConvertTimeToUtc(ByRef DateTime As DateTime) As DateTime
Attribute ConvertTimeToUtc.VB_Description = "Converts the specified date and time to Coordinated Universal Time (UTC)."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeToUtc(DateTime.ComObject)
      Set ConvertTimeToUtc = .Self
   End With
End Function

'@Static
'@Description("Converts the time in a specified time zone to Coordinated Universal Time (UTC).")
Public Function ConvertTimeToUtc2(ByRef DateTime As DateTime, ByRef sourceTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTimeToUtc2.VB_Description = "Converts the time in a specified time zone to Coordinated Universal Time (UTC)."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeToUtc2(DateTime.ComObject, sourceTimeZone.ComObject)
      Set ConvertTimeToUtc2 = .Self
   End With
End Function

'@Static
'@Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name.")
Public Function CreateCustomTimeZone(ByVal Id As String, ByRef BaseUtcOffset As TimeSpan, ByVal DisplayName As String, ByVal standardDisplayName As String) As TimeZoneInfo
Attribute CreateCustomTimeZone.VB_Description = "Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name."
   With New TimeZoneInfo
      Set .ComObject = this.DotNetLibTimeZoneInfo.CreateCustomTimeZone(Id, BaseUtcOffset.ComObject, DisplayName, standardDisplayName)
      Set CreateCustomTimeZone = .Self
   End With
End Function

'@Description("Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal.")
Public Function Equals(ByRef other As TimeZoneInfo) As Boolean
Attribute Equals.VB_Description = "Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal."
   Equals = this.DotNetLibTimeZoneInfo.Equals(other.ComObject)
End Function

'@Description("Determines whether the current TimeZoneInfo object and another object are equal.")
Public Function Equals2(ByVal obj As Variant) As Boolean
Attribute Equals2.VB_Description = "Determines whether the current TimeZoneInfo object and another object are equal."
   Equals2 = this.DotNetLibTimeZoneInfo.Equals2(obj)
End Function

'@Static
'@Description("Instantiates a new TimeZoneInfo object based on its identifier.")
Public Function FindSystemTimeZoneById(ByVal Id As String) As TimeZoneInfo
Attribute FindSystemTimeZoneById.VB_Description = "Instantiates a new TimeZoneInfo object based on its identifier."
   With New TimeZoneInfo
      Set .ComObject = this.DotNetLibTimeZoneInfo.FindSystemTimeZoneById(Id)
      Set FindSystemTimeZoneById = .Self
   End With
End Function

'@Static
'@Description("Deserializes a string to re-create an original serialized TimeZoneInfo object.")
Public Function FromSerializedString(ByVal source As String) As TimeZoneInfo
Attribute FromSerializedString.VB_Description = "Deserializes a string to re-create an original serialized TimeZoneInfo object."
   With New TimeZoneInfo
      Set .ComObject = this.DotNetLibTimeZoneInfo.FromSerializedString(source)
      Set FromSerializedString = .Self
   End With
End Function

'@TODO Require to convert to an array of wrapped  DotNetLib.TimeSpan
'@Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")
Public Function GetAmbiguousTimeOffsets(ByRef DateTime As DateTime) As DotNetLib.TimeSpan()
Attribute GetAmbiguousTimeOffsets.VB_Description = "Returns information about the possible dates and times that an ambiguous date and time can be mapped to."
   GetAmbiguousTimeOffsets = this.DotNetLibTimeZoneInfo.GetAmbiguousTimeOffsets(DateTime.ComObject)
End Function

'@TODO Require to convert to an array of wrapped  DotNetLib.TimeSpan
'@Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")
Public Function GetAmbiguousTimeOffsets2(ByRef DateTimeOffset As DateTimeOffset) As DotNetLib.TimeSpan()
Attribute GetAmbiguousTimeOffsets2.VB_Description = "Returns information about the possible dates and times that an ambiguous date and time can be mapped to."
   GetAmbiguousTimeOffsets2 = this.DotNetLibTimeZoneInfo.GetAmbiguousTimeOffsets2(DateTimeOffset.ComObject)
End Function

'@Description("Serves as a hash function for hashing algorithms and data structures such as hash tables.")
Public Function GetHashCode() As Long
Attribute GetHashCode.VB_Description = "Serves as a hash function for hashing algorithms and data structures such as hash tables."
   GetHashCode = this.DotNetLibTimeZoneInfo.GetHashCode()
End Function

'@Static
'@Description("Returns a sorted collection of all the time zones about which information is available on the local system.")
Public Function GetSystemTimeZones() As DotNetLib.ReadOnlyCollection
Attribute GetSystemTimeZones.VB_Description = "Returns a sorted collection of all the time zones about which information is available on the local system."
   Set GetSystemTimeZones = this.DotNetLibTimeZoneInfo.GetSystemTimeZones()
End Function

'@Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")
Public Function GetUtcOffset(ByRef DateTime As DateTime) As TimeSpan
Attribute GetUtcOffset.VB_Description = "Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time."
   With New TimeSpan
      Set .ComObject = this.DotNetLibTimeZoneInfo.GetUtcOffset(DateTime.ComObject)
      Set GetUtcOffset = .Self
   End With
End Function

'@Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")
Public Function GetUtcOffset2(ByRef DateTimeOffset As DateTimeOffset) As TimeSpan
Attribute GetUtcOffset2.VB_Description = "Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time."
   With New TimeSpan
      Set .ComObject = this.DotNetLibTimeZoneInfo.GetUtcOffset2(DateTimeOffset.ComObject)
      Set GetUtcOffset2 = .Self
   End With
End Function

'@Description("Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules.")
Public Function HasSameRules(ByRef other As TimeZoneInfo) As Boolean
Attribute HasSameRules.VB_Description = "Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules."
   HasSameRules = this.DotNetLibTimeZoneInfo.HasSameRules(other.ComObject)
End Function

'@Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")
Public Function IsAmbiguousTime(ByRef DateTime As DateTime) As Boolean
Attribute IsAmbiguousTime.VB_Description = "Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times."
   IsAmbiguousTime = this.DotNetLibTimeZoneInfo.IsAmbiguousTime(DateTime.ComObject)
End Function

'@Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")
Public Function IsAmbiguousTime2(ByRef DateTimeOffset As DateTimeOffset) As Boolean
Attribute IsAmbiguousTime2.VB_Description = "Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times."
   IsAmbiguousTime2 = this.DotNetLibTimeZoneInfo.IsAmbiguousTime2(DateTimeOffset.ComObject)
End Function

'@Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")
Public Function IsDaylightSavingTime(ByRef DateTime As DateTime) As Boolean
Attribute IsDaylightSavingTime.VB_Description = "Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object."
   IsDaylightSavingTime = this.DotNetLibTimeZoneInfo.IsDaylightSavingTime(DateTime.ComObject)
End Function

'@Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")
Public Function IsDaylightSavingTime2(ByRef DateTimeOffset As DateTimeOffset) As Boolean
Attribute IsDaylightSavingTime2.VB_Description = "Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object."
   IsDaylightSavingTime2 = this.DotNetLibTimeZoneInfo.IsDaylightSavingTime2(DateTimeOffset.ComObject)
End Function

'@Description("Indicates whether a particular date and time is invalid.")
Public Function IsInvalidTime(ByRef DateTime As DateTime) As Boolean
Attribute IsInvalidTime.VB_Description = "Indicates whether a particular date and time is invalid."
   IsInvalidTime = this.DotNetLibTimeZoneInfo.IsInvalidTime(DateTime.ComObject)
End Function

'@Description("Converts the current TimeZoneInfo object to a serialized string.")
Public Function ToSerializedString() As String
Attribute ToSerializedString.VB_Description = "Converts the current TimeZoneInfo object to a serialized string."
   ToSerializedString = this.DotNetLibTimeZoneInfo.ToSerializedString()
End Function

'@DefaultMember
Public Property Get ToString() As String
Attribute ToString.VB_UserMemId = 0
   ToString = this.DotNetLibTimeZoneInfo.ToString()
End Property
