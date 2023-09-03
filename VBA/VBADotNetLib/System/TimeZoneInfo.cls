VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeZoneInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "TimeZoneInfo factory methods and static members."
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId
'@ModuleDescription("TimeZoneInfo factory methods and static members.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 August 4, 2023
'@LastModified September 3, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timezoneinfo?view=netframework-4.8.1

'@Remarks
' Singleton class for TimeZoneInfo containing constructors and static members.
' Renamed DotNetlib.TimeZoneInfo.Local to Locale due to VBA reserved word

Option Explicit

Private Type TTimeZoneInfo
   TZISingleton As DotNetLib.TimeZoneInfoSingleton
End Type

Private this As TTimeZoneInfo

Private Sub Class_Initialize()
   Set this.TZISingleton = New DotNetLib.TimeZoneInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.TZISingleton = Nothing
End Sub

'--------------------------------------------------------------------
'Properties
'--------------------------------------------------------------------

'@Static
'@Description("Gets a TimeZoneInfo object that represents the local time zone.")
'@Note Rename Local to Locale due to VBA resevered word
Public Property Get Locale() As DotNetLib.TimeZoneInfo
Attribute Locale.VB_Description = "Gets a TimeZoneInfo object that represents the local time zone."
    Set Locale = this.TZISingleton.Local
End Property

'@Static
'@Description("Gets a TimeZoneInfo object that represents the Coordinated Universal Time (UTC) zone.")
Public Property Get Utc() As DotNetLib.TimeZoneInfo
Attribute Utc.VB_Description = "Gets a TimeZoneInfo object that represents the Coordinated Universal Time (UTC) zone."
    Set Utc = this.TZISingleton.Utc()
End Property

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Static
'@Description("Clears cached time zone data.")
Public Sub ClearCachedData()
Attribute ClearCachedData.VB_Description = "Clears cached time zone data."
   this.TZISingleton.ClearCachedData
End Sub

'@Static
'@Description("Converts a time to the time in a particular time zone.")
Public Function ConvertTime(ByVal sourceDateTime As DotNetLib.DateTime, ByVal destinationTimeZone As DotNetLib.TimeZoneInfo) As DotNetLib.DateTime
Attribute ConvertTime.VB_Description = "Converts a time to the time in a particular time zone."
    Set ConvertTime = this.TZISingleton.ConvertTime(sourceDateTime, destinationTimeZone)
End Function

'@Static
'@Description("Converts a time to the time in a particular time zone.")
Public Function ConvertTime2(ByVal sourceDateTimeOffset As DotNetLib.DateTimeOffset, ByVal destinationTimeZone As DotNetLib.TimeZoneInfo) As DotNetLib.DateTimeOffset
Attribute ConvertTime2.VB_Description = "Converts a time to the time in a particular time zone."
    Set ConvertTime2 = this.TZISingleton.ConvertTime2(sourceDateTimeOffset, destinationTimeZone)
End Function

'@Static
'@Description("Converts a time from one time zone to another.")
Public Function ConvertTime3(ByVal sourceDateTime As DotNetLib.DateTime, ByVal sourceTimeZone As DotNetLib.TimeZoneInfo, ByVal destinationTimeZone As DotNetLib.TimeZoneInfo) As DotNetLib.DateTime
Attribute ConvertTime3.VB_Description = "Converts a time from one time zone to another."
    Set ConvertTime3 = this.TZISingleton.ConvertTime3(sourceDateTime, sourceTimeZone, destinationTimeZone)
End Function

'@Static
'@Description("Converts a time to the time in another time zone based on the time zone's identifier.")
Public Function ConvertTimeBySystemTimeZoneId(ByVal sourceDateTime As DotNetLib.DateTime, ByVal destinationTimeZoneId As String) As DotNetLib.DateTime
Attribute ConvertTimeBySystemTimeZoneId.VB_Description = "Converts a time to the time in another time zone based on the time zone's identifier."
    Set ConvertTimeBySystemTimeZoneId = this.TZISingleton.ConvertTimeBySystemTimeZoneId(sourceDateTime, destinationTimeZoneId)
End Function

'@Static
'@Description("Converts a time to the time in another time zone based on the time zone's identifier.")
Public Function ConvertTimeBySystemTimeZoneId2(ByVal sourceDateTimeOffset As DotNetLib.DateTimeOffset, ByVal destinationTimeZoneId As String) As DotNetLib.DateTimeOffset
Attribute ConvertTimeBySystemTimeZoneId2.VB_Description = "Converts a time to the time in another time zone based on the time zone's identifier."
    Set ConvertTimeBySystemTimeZoneId2 = this.TZISingleton.ConvertTimeBySystemTimeZoneId2(sourceDateTimeOffset, destinationTimeZoneId)
End Function

'@Static
'@Description("Converts a time from one time zone to another based on time zone identifiers.")
Public Function ConvertTimeBySystemTimeZoneId3(ByVal sourceDateTime As DotNetLib.DateTime, ByVal sourceTimeZoneId As String, ByVal destinationTimeZoneId As String) As DotNetLib.DateTime
Attribute ConvertTimeBySystemTimeZoneId3.VB_Description = "Converts a time from one time zone to another based on time zone identifiers."
    Set ConvertTimeBySystemTimeZoneId3 = this.TZISingleton.ConvertTimeBySystemTimeZoneId3(sourceDateTime, sourceTimeZoneId, destinationTimeZoneId)
End Function

'@Static
'@Description("Converts a Coordinated Universal Time (UTC) to the time in a specified time zone.")
Public Function ConvertTimeFromUtc(ByVal sourceDateTime As DotNetLib.DateTime, ByVal destinationTimeZone As DotNetLib.TimeZoneInfo) As DotNetLib.DateTime
Attribute ConvertTimeFromUtc.VB_Description = "Converts a Coordinated Universal Time (UTC) to the time in a specified time zone."
    Set ConvertTimeFromUtc = this.TZISingleton.ConvertTimeFromUtc(sourceDateTime, destinationTimeZone)
End Function

'@Static
'@Description("Converts the specified date and time to Coordinated Universal Time (UTC).")
Public Function ConvertTimeToUtc(ByVal sourceDateTime As DotNetLib.DateTime) As DotNetLib.DateTime
Attribute ConvertTimeToUtc.VB_Description = "Converts the specified date and time to Coordinated Universal Time (UTC)."
    Set ConvertTimeToUtc = this.TZISingleton.ConvertTimeToUtc(sourceDateTime)
End Function

'@Static
'@Description("Converts the time in a specified time zone to Coordinated Universal Time (UTC).")
Public Function ConvertTimeToUtc2(ByVal sourceDateTime As DotNetLib.DateTime, ByVal sourceTimeZone As DotNetLib.TimeZoneInfo) As DotNetLib.DateTime
Attribute ConvertTimeToUtc2.VB_Description = "Converts the time in a specified time zone to Coordinated Universal Time (UTC)."
    Set ConvertTimeToUtc2 = this.TZISingleton.ConvertTimeToUtc2(sourceDateTime, sourceTimeZone)
End Function

'@Static
'@Description("Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name.")
Public Function CreateCustomTimeZone(ByVal Id As String, ByVal BaseUtcOffset As DotNetLib.TimeSpan, ByVal DisplayName As String, ByVal standardDisplayName As String) As DotNetLib.TimeZoneInfo
Attribute CreateCustomTimeZone.VB_Description = "Creates a custom time zone with a specified identifier, an offset from Coordinated Universal Time (UTC), a display name, and a standard time display name."
    Set CreateCustomTimeZone = this.TZISingleton.CreateCustomTimeZone(Id, BaseUtcOffset, DisplayName, standardDisplayName)
End Function

'@Static
'@Description("Instantiates a new TimeZoneInfo object based on its identifier.")
Public Function FindSystemTimeZoneById(ByVal Id As String) As DotNetLib.TimeZoneInfo
Attribute FindSystemTimeZoneById.VB_Description = "Instantiates a new TimeZoneInfo object based on its identifier."
    Set FindSystemTimeZoneById = this.TZISingleton.FindSystemTimeZoneById(Id)
End Function

'@Static
'@Description("Deserializes a string to re-create an original serialized TimeZoneInfo object.")
Public Function FromSerializedString(ByVal Source As String) As DotNetLib.TimeZoneInfo
Attribute FromSerializedString.VB_Description = "Deserializes a string to re-create an original serialized TimeZoneInfo object."
    Set FromSerializedString = this.TZISingleton.FromSerializedString(Source)
End Function

'@Static
'@Description("Returns a sorted collection of all the time zones about which information is available on the local system.")
Public Function GetSystemTimeZones() As DotNetLib.ReadOnlyCollection
Attribute GetSystemTimeZones.VB_Description = "Returns a sorted collection of all the time zones about which information is available on the local system."
   Set GetSystemTimeZones = this.TZISingleton.GetSystemTimeZones()
End Function
