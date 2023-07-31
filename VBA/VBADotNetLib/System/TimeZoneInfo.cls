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
'@LastModified July 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1

'@Remarks
' Renamed DotNetlib.TimeZoneInfo.Local to Locale due to VBA reserved word

Option Explicit

Implements ITimeZoneInfo

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

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------

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

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Description("Clears cached time zone data.")
Public Sub ClearCachedData()
Attribute ClearCachedData.VB_Description = "Clears cached time zone data."
   this.DotNetLibTimeZoneInfo.ClearCachedData
End Sub

'@Static
'@Description("Converts a time to the time in a particular time zone.")
Public Function ConvertTime(ByVal sourceDateTime As DateTime, ByVal destinationTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTime.VB_Description = "Converts a time to the time in a particular time zone."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTime(sourceDateTime.ComObject, destinationTimeZone.ComObject)
      Set ConvertTime = .Self
   End With
End Function

'@Static
'@Description("Converts a time to the time in a particular time zone.")
Public Function ConvertTime2(ByVal sourceDateTimeOffset As DateTimeOffset, ByVal destinationTimeZone As TimeZoneInfo) As DateTimeOffset
Attribute ConvertTime2.VB_Description = "Converts a time to the time in a particular time zone."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTime2(sourceDateTimeOffset.ComObject, destinationTimeZone.ComObject)
      Set ConvertTime2 = .Self
   End With
End Function

'@Static
'@Description("Converts a time from one time zone to another.")
Public Function ConvertTime3(ByVal sourceDateTime As DateTime, ByVal sourceTimeZone As TimeZoneInfo, ByVal destinationTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTime3.VB_Description = "Converts a time from one time zone to another."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTime3(sourceDateTime.ComObject, sourceTimeZone.ComObject, destinationTimeZone.ComObject)
      Set ConvertTime3 = .Self
   End With
End Function

'@Static
'@Description("Converts a time to the time in another time zone based on the time zone's identifier.")
Public Function ConvertTimeBySystemTimeZoneId(ByVal sourceDateTime As DateTime, ByVal destinationTimeZoneId As String) As DateTime
Attribute ConvertTimeBySystemTimeZoneId.VB_Description = "Converts a time to the time in another time zone based on the time zone's identifier."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeBySystemTimeZoneId(sourceDateTime.ComObject, destinationTimeZoneId)
      Set ConvertTimeBySystemTimeZoneId = .Self
   End With
End Function

'@Static
'@Description("Converts a time to the time in another time zone based on the time zone's identifier.")
Public Function ConvertTimeBySystemTimeZoneId2(ByVal sourceDateTimeOffset As DateTimeOffset, ByVal destinationTimeZoneId As String) As DateTimeOffset
Attribute ConvertTimeBySystemTimeZoneId2.VB_Description = "Converts a time to the time in another time zone based on the time zone's identifier."
   With New DateTimeOffset
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeBySystemTimeZoneId2(sourceDateTimeOffset.ComObject, destinationTimeZoneId)
      Set ConvertTimeBySystemTimeZoneId2 = .Self
   End With
End Function

'@Static
'@Description("Converts a time from one time zone to another based on time zone identifiers.")
Public Function ConvertTimeBySystemTimeZoneId3(ByVal sourceDateTime As DateTime, ByVal sourceTimeZoneId As String, ByVal destinationTimeZoneId As String) As DateTime
Attribute ConvertTimeBySystemTimeZoneId3.VB_Description = "Converts a time from one time zone to another based on time zone identifiers."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeBySystemTimeZoneId3(sourceDateTime.ComObject, sourceTimeZoneId, destinationTimeZoneId)
      Set ConvertTimeBySystemTimeZoneId3 = .Self
   End With
End Function

'@Static
'@Description("Converts a Coordinated Universal Time (UTC) to the time in a specified time zone.")
Public Function ConvertTimeFromUtc(ByVal sourceDateTime As DateTime, ByVal destinationTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTimeFromUtc.VB_Description = "Converts a Coordinated Universal Time (UTC) to the time in a specified time zone."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeFromUtc(sourceDateTime.ComObject, destinationTimeZone.ComObject)
      Set ConvertTimeFromUtc = .Self
   End With
End Function

'@Static
'@Description("Converts the specified date and time to Coordinated Universal Time (UTC).")
Public Function ConvertTimeToUtc(ByVal sourceDateTime As DateTime) As DateTime
Attribute ConvertTimeToUtc.VB_Description = "Converts the specified date and time to Coordinated Universal Time (UTC)."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeToUtc(sourceDateTime.ComObject)
      Set ConvertTimeToUtc = .Self
   End With
End Function

'@Static
'@Description("Converts the time in a specified time zone to Coordinated Universal Time (UTC).")
Public Function ConvertTimeToUtc2(ByVal sourceDateTime As DateTime, ByVal sourceTimeZone As TimeZoneInfo) As DateTime
Attribute ConvertTimeToUtc2.VB_Description = "Converts the time in a specified time zone to Coordinated Universal Time (UTC)."
   With New DateTime
      Set .ComObject = this.DotNetLibTimeZoneInfo.ConvertTimeToUtc2(sourceDateTime.ComObject, sourceTimeZone.ComObject)
      Set ConvertTimeToUtc2 = .Self
   End With
End Function

'@Static
'@Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name.")
Public Function CreateCustomTimeZone(ByVal Id As String, ByVal BaseUtcOffset As TimeSpan, ByVal DisplayName As String, ByVal standardDisplayName As String) As TimeZoneInfo
Attribute CreateCustomTimeZone.VB_Description = "Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name."
   With New TimeZoneInfo
      Set .ComObject = this.DotNetLibTimeZoneInfo.CreateCustomTimeZone(Id, BaseUtcOffset.ComObject, DisplayName, standardDisplayName)
      Set CreateCustomTimeZone = .Self
   End With
End Function

'@Description("Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal.")
Public Function Equals(ByVal other As TimeZoneInfo) As Boolean
Attribute Equals.VB_Description = "Determines whether the current TimeZoneInfo object and another TimeZoneInfo object are equal."
   Equals = this.DotNetLibTimeZoneInfo.Equals(other.ComObject)
End Function

'@Description("Determines whether the current TimeZoneInfo object and another object are equal.")
Public Function Equals2(ByVal obj As Variant) As Boolean
Attribute Equals2.VB_Description = "Determines whether the current TimeZoneInfo object and another object are equal."
    If Not obj Is Nothing Then
        If TypeOf obj Is TimeZoneInfo Then
           Dim tzi As TimeZoneInfo
           Set tzi = obj
           Equals2 = this.DotNetLibTimeZoneInfo.Equals(tzi.ComObject)
        End If
    End If
End Function

'@Static
'@Description("Instantiates a new TimeZoneInfo object based on its identifier.")
Public Function FindSystemTimeZoneById(ByVal Id As String) As ITimeZoneInfo
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

'@Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")
Public Function GetAmbiguousTimeOffsets(ByVal sourceDateTime As DateTime) As TimeSpan()
Attribute GetAmbiguousTimeOffsets.VB_Description = "Returns information about the possible dates and times that an ambiguous date and time can be mapped to."
   Dim ambiguousTimeOffsets() As DotNetLib.TimeSpan
   ambiguousTimeOffsets = this.DotNetLibTimeZoneInfo.GetAmbiguousTimeOffsets(sourceDateTime.ComObject)
   
   Dim output() As TimeSpan
   ReDim output(UBound(ambiguousTimeOffsets))
   Dim arrayIndex As Long
   For arrayIndex = LBound(ambiguousTimeOffsets) To UBound(ambiguousTimeOffsets)
      With New TimeSpan
         Set .ComObject = ambiguousTimeOffsets(arrayIndex)
         Set output(arrayIndex) = .Self
      End With
   Next arrayIndex
   GetAmbiguousTimeOffsets = output
End Function

'@Description("Returns information about the possible dates and times that an ambiguous date and time can be mapped to.")
Public Function GetAmbiguousTimeOffsets2(ByVal sourceDateTimeOffset As DateTimeOffset) As TimeSpan()
Attribute GetAmbiguousTimeOffsets2.VB_Description = "Returns information about the possible dates and times that an ambiguous date and time can be mapped to."
   Dim ambiguousTimeOffsets() As DotNetLib.TimeSpan
   ambiguousTimeOffsets = this.DotNetLibTimeZoneInfo.GetAmbiguousTimeOffsets2(sourceDateTimeOffset.ComObject)
   
   Dim output() As TimeSpan
   ReDim output(UBound(ambiguousTimeOffsets))
   Dim arrayIndex As Long
   For arrayIndex = LBound(ambiguousTimeOffsets) To UBound(ambiguousTimeOffsets)
      With New TimeSpan
         Set .ComObject = ambiguousTimeOffsets(arrayIndex)
         Set output(arrayIndex) = .Self
      End With
   Next arrayIndex
   GetAmbiguousTimeOffsets2 = output
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
Public Function GetUtcOffset(ByVal sourceDateTime As DateTime) As TimeSpan
Attribute GetUtcOffset.VB_Description = "Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time."
   With New TimeSpan
      Set .ComObject = this.DotNetLibTimeZoneInfo.GetUtcOffset(sourceDateTime.ComObject)
      Set GetUtcOffset = .Self
   End With
End Function

'@Description("Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time.")
Public Function GetUtcOffset2(ByVal sourceDateTimeOffset As DateTimeOffset) As TimeSpan
Attribute GetUtcOffset2.VB_Description = "Calculates the offset or difference between the time in this time zone and Coordinated Universal Time (UTC) for a particular date and time."
   With New TimeSpan
      Set .ComObject = this.DotNetLibTimeZoneInfo.GetUtcOffset2(sourceDateTimeOffset.ComObject)
      Set GetUtcOffset2 = .Self
   End With
End Function

'@Description("Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules.")
Public Function HasSameRules(ByVal other As TimeZoneInfo) As Boolean
Attribute HasSameRules.VB_Description = "Indicates whether the current object and another TimeZoneInfo object have the same adjustment rules."
   HasSameRules = this.DotNetLibTimeZoneInfo.HasSameRules(other.ComObject)
End Function

'@Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")
Public Function IsAmbiguousTime(ByVal sourceDateTime As DateTime) As Boolean
Attribute IsAmbiguousTime.VB_Description = "Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times."
   IsAmbiguousTime = this.DotNetLibTimeZoneInfo.IsAmbiguousTime(sourceDateTime.ComObject)
End Function

'@Description("Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times.")
Public Function IsAmbiguousTime2(ByVal sourceDateTimeOffset As DateTimeOffset) As Boolean
Attribute IsAmbiguousTime2.VB_Description = "Determines whether a particular date and time in a particular time zone is ambiguous and can be mapped to two or more Coordinated Universal Time (UTC) times."
   IsAmbiguousTime2 = this.DotNetLibTimeZoneInfo.IsAmbiguousTime2(sourceDateTimeOffset.ComObject)
End Function

'@Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")
Public Function IsDaylightSavingTime(ByVal sourceDateTime As DateTime) As Boolean
Attribute IsDaylightSavingTime.VB_Description = "Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object."
   IsDaylightSavingTime = this.DotNetLibTimeZoneInfo.IsDaylightSavingTime(sourceDateTime.ComObject)
End Function

'@Description("Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object.")
Public Function IsDaylightSavingTime2(ByVal sourceDateTimeOffset As DateTimeOffset) As Boolean
Attribute IsDaylightSavingTime2.VB_Description = "Indicates whether a specified date and time falls in the range of daylight saving time for the time zone of the current TimeZoneInfo object."
   IsDaylightSavingTime2 = this.DotNetLibTimeZoneInfo.IsDaylightSavingTime2(sourceDateTimeOffset.ComObject)
End Function

'@Description("Indicates whether a particular date and time is invalid.")
Public Function IsInvalidTime(ByVal sourceDateTime As DateTime) As Boolean
Attribute IsInvalidTime.VB_Description = "Indicates whether a particular date and time is invalid."
   IsInvalidTime = this.DotNetLibTimeZoneInfo.IsInvalidTime(sourceDateTime.ComObject)
End Function

'@Description("Converts the current TimeZoneInfo object to a serialized string.")
Public Function ToSerializedString() As String
Attribute ToSerializedString.VB_Description = "Converts the current TimeZoneInfo object to a serialized string."
   ToSerializedString = this.DotNetLibTimeZoneInfo.ToSerializedString()
End Function

'@DefaultMember
'@Description("Returns the current TimeZoneInfo object's display name.")
Public Property Get ToString() As String
Attribute ToString.VB_Description = "Returns the current TimeZoneInfo object's display name."
Attribute ToString.VB_UserMemId = 0
   ToString = this.DotNetLibTimeZoneInfo.ToString()
End Property

'--------------------------------------------------------------------
'Interface ITimeZoneInfo
'--------------------------------------------------------------------

Private Property Get ITimeZoneInfo_BaseUtcOffset() As TimeSpan
    Set ITimeZoneInfo_BaseUtcOffset = BaseUtcOffset
End Property

Private Property Get ITimeZoneInfo_DaylightName() As String
    ITimeZoneInfo_DaylightName = DaylightName
End Property

Private Property Get ITimeZoneInfo_DisplayName() As String
    ITimeZoneInfo_DisplayName = DisplayName
End Property

Private Property Get ITimeZoneInfo_Id() As String
    ITimeZoneInfo_Id = Id
End Property

Private Property Get ITimeZoneInfo_StandardName() As String
    ITimeZoneInfo_StandardName = StandardName
End Property

Private Property Get ITimeZoneInfo_SupportsDaylightSavingTime() As Boolean
    ITimeZoneInfo_SupportsDaylightSavingTime = SupportsDaylightSavingTime
End Property

Private Function ITimeZoneInfo_Equals(ByVal other As TimeZoneInfo) As Boolean
    ITimeZoneInfo_Equals = Equals(other)
End Function

'@Description("Determines whether the current TimeZoneInfo object and another object are equal.")
Public Function ITimeZoneInfo_Equals2(ByVal obj As Variant) As Boolean
Attribute ITimeZoneInfo_Equals2.VB_Description = "Determines whether the current TimeZoneInfo object and another object are equal."
    ITimeZoneInfo_Equals2 = Equals2(obj)
End Function

Private Function ITimeZoneInfo_GetAmbiguousTimeOffsets(ByVal sourceDateTime As DateTime) As TimeSpan()
    ITimeZoneInfo_GetAmbiguousTimeOffsets = GetAmbiguousTimeOffsets(sourceDateTime)
End Function

Private Function ITimeZoneInfo_GetAmbiguousTimeOffsets2(ByVal sourceDateTimeOffset As DateTimeOffset) As TimeSpan()
    ITimeZoneInfo_GetAmbiguousTimeOffsets2 = GetAmbiguousTimeOffsets2(sourceDateTimeOffset)
End Function

Private Function ITimeZoneInfo_GetHashCode() As Long
    ITimeZoneInfo_GetHashCode = GetHashCode()
End Function

Private Function ITimeZoneInfo_GetUtcOffset(ByVal sourceDateTime As DateTime) As TimeSpan
    Set ITimeZoneInfo_GetUtcOffset = GetUtcOffset(sourceDateTime)
End Function

Private Function ITimeZoneInfo_GetUtcOffset2(ByVal sourceDateTimeOffset As DateTimeOffset) As TimeSpan
    Set ITimeZoneInfo_GetUtcOffset2 = GetUtcOffset2(sourceDateTimeOffset)
End Function

Private Function ITimeZoneInfo_HasSameRules(ByVal other As TimeZoneInfo) As Boolean
    ITimeZoneInfo_HasSameRules = HasSameRules(other)
End Function

Private Function ITimeZoneInfo_IsAmbiguousTime(ByVal sourceDateTime As DateTime) As Boolean
    ITimeZoneInfo_IsAmbiguousTime = IsAmbiguousTime(sourceDateTime)
End Function

Private Function ITimeZoneInfo_IsAmbiguousTime2(ByVal sourceDateTimeOffset As DateTimeOffset) As Boolean
    ITimeZoneInfo_IsAmbiguousTime2 = IsAmbiguousTime2(sourceDateTimeOffset)
End Function

Private Function ITimeZoneInfo_IsDaylightSavingTime(ByVal sourceDateTime As DateTime) As Boolean
    ITimeZoneInfo_IsDaylightSavingTime = IsDaylightSavingTime(sourceDateTime)
End Function

Private Function ITimeZoneInfo_IsDaylightSavingTime2(ByVal sourceDateTimeOffset As DateTimeOffset) As Boolean
    ITimeZoneInfo_IsDaylightSavingTime2 = IsDaylightSavingTime2(sourceDateTimeOffset)
End Function

Private Function ITimeZoneInfo_IsInvalidTime(ByVal sourceDateTime As DateTime) As Boolean
    ITimeZoneInfo_IsInvalidTime = IsInvalidTime(sourceDateTime)
End Function

Private Function ITimeZoneInfo_ToSerializedString() As String
    ITimeZoneInfo_ToSerializedString = ToSerializedString()
End Function

Private Property Get ITimeZoneInfo_ToString() As String
    ITimeZoneInfo_ToString = ToString
End Property
