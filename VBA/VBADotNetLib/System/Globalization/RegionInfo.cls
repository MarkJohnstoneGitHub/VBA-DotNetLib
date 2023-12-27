VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Contains information about the country/region."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Contains information about the country/region.")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo?view=netframework-4.8.1

'@Remarks
' Singleton class for RegionInfo containing constructors and static members.

Option Explicit

Private Type TRegionInfoSingleton
   RegionInfoSingleton As DotNetLib.RegionInfoSingleton
End Type

Private this As TRegionInfoSingleton

Private Sub Class_Initialize()
   Set this.RegionInfoSingleton = New DotNetLib.RegionInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.RegionInfoSingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the RegionInfo class based on the country/region associated with the specified culture identifier.")
Public Function Create(ByVal culture As Long) As DotNetLib.RegionInfo
Attribute Create.VB_Description = "Initializes a new instance of the RegionInfo class based on the country/region associated with the specified culture identifier."
   Set Create = this.RegionInfoSingleton.Create(culture)
End Function

'@Static
'@Description("Initializes a new instance of the RegionInfo class based on the country/region or specific culture, specified by name.")
Public Function Create2(ByVal Name As String) As DotNetLib.RegionInfo
Attribute Create2.VB_Description = "Initializes a new instance of the RegionInfo class based on the country/region or specific culture, specified by name."
   Set Create2 = this.RegionInfoSingleton.Create2(Name)
End Function

'@Static
'@Description("Gets the RegionInfo that represents the country/region used by the current thread.")
Public Property Get CurrentRegion() As DotNetLib.RegionInfo
Attribute CurrentRegion.VB_Description = "Gets the RegionInfo that represents the country/region used by the current thread."
   Set CurrentRegion = this.RegionInfoSingleton.CurrentRegion()
End Property
