VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CelestialInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "CelestialInfo factory methods and static members."
'@Folder "CoordinateSharp.CelestialInfo"
'@PredeclaredId
'@ModuleDescription("CelestialInfo factory methods and static members.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 August 19, 2023
'@LastModified August 19, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://coordinatesharp.com/Help/html/T_CoordinateSharp_Celestial.htm

Option Explicit

Private Type TCelestialInfo
   DotNetLibCelestialInfo As DotNetLib.CelestialInfo
End Type

Private this As TCelestialInfo

Private Sub Class_Initialize()
   Set this.DotNetLibCelestialInfo = New DotNetLib.CelestialInfo
End Sub

Private Sub Class_Terminate()
   Set this.DotNetLibCelestialInfo = Nothing
End Sub

'@Description("Calcualtes sunrise time.")
Public Function SunRise(ByVal latitude As Double, ByVal longitude As Double, ByVal pDateTime As DotNetLib.DateTime) As DotNetLib.DateTime
Attribute SunRise.VB_Description = "Calcualtes sunrise time."
   Set SunRise = this.DotNetLibCelestialInfo.SunRise(latitude, longitude, pDateTime)
End Function

'@Description("Calculates sunset time.")
Public Function SunSet(ByVal latitude As Double, ByVal longitude As Double, ByRef pDateTime As DotNetLib.DateTime) As DotNetLib.DateTime
Attribute SunSet.VB_Description = "Calculates sunset time."
   Set SunSet = this.DotNetLibCelestialInfo.SunSet(latitude, longitude, pDateTime)
End Function
