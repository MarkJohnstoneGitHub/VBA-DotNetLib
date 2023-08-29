VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "DateTimeFormatInfo factory methods and static members that provides culture-specific information about the format of date and time values."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("DateTimeFormatInfo factory methods and static members that provides culture-specific information about the format of date and time values.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 29, 2023
'@LastModified August 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo?view=netframework-4.8.1

Option Explicit

Private Type TDateTimeFormatInfo
   DotNetLibDateTimeFormatInfoSingleton As DotNetLib.DateTimeFormatInfoSingleton
End Type

Private this As TDateTimeFormatInfo

Private Sub Class_Initialize()
   Set this.DotNetLibDateTimeFormatInfoSingleton = New DotNetLib.DateTimeFormatInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.DotNetLibDateTimeFormatInfoSingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new writable instance of the DateTimeFormatInfo class that is culture-independent (invariant).")
Public Function Create() As DotNetLib.DateTimeFormatInfo
Attribute Create.VB_Description = "Initializes a new writable instance of the DateTimeFormatInfo class that is culture-independent (invariant)."
   Set Create = this.DotNetLibDateTimeFormatInfoSingleton.Create()
End Function

'@Static
'@Description("Gets a read-only DateTimeFormatInfo object that formats values based on the current culture.")
Public Property Get CurrentInfo() As DotNetLib.DateTimeFormatInfo
Attribute CurrentInfo.VB_Description = "Gets a read-only DateTimeFormatInfo object that formats values based on the current culture."
   Set CurrentInfo = this.DotNetLibDateTimeFormatInfoSingleton.CurrentInfo()
End Property

'@Static
'@Description("Gets the default read-only DateTimeFormatInfo object that is culture-independent (invariant).")
Public Property Get InvariantInfo() As DotNetLib.DateTimeFormatInfo
Attribute InvariantInfo.VB_Description = "Gets the default read-only DateTimeFormatInfo object that is culture-independent (invariant)."
   Set InvariantInfo = this.DotNetLibDateTimeFormatInfoSingleton.InvariantInfo()
End Property

'@Static
'@Description("Returns the DateTimeFormatInfo object associated with the specified IFormatProvider.")
Public Function GetInstance(ByRef formatProvider As IFormatProvider) As DotNetLib.DateTimeFormatInfo
Attribute GetInstance.VB_Description = "Returns the DateTimeFormatInfo object associated with the specified IFormatProvider."
   Set GetInstance = this.DotNetLibDateTimeFormatInfoSingleton.GetInstance(formatProvider)
End Function

'@Static
'@Description("Returns a read-only DateTimeFormatInfo wrapper.")
Public Function ReadOnly(ByRef dtfi As DotNetLib.DateTimeFormatInfo) As DotNetLib.DateTimeFormatInfo
Attribute ReadOnly.VB_Description = "Returns a read-only DateTimeFormatInfo wrapper."
   Set ReadOnly = this.DotNetLibDateTimeFormatInfoSingleton.ReadOnly(dtfi)
End Function
