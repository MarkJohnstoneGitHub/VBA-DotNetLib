VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NumberFormatInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "NumberFormatInfo factory methods and static members that provide culture-specific information for formatting and parsing numeric values."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("NumberFormatInfo factory methods and static members that provide culture-specific information for formatting and parsing numeric values.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 28, 2023
'@LastModified August 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.numberformatinfo?view=netframework-4.8.1

Option Explicit

Private Type TNumberFormatInfoSingleton
   NumberFormatInfoSingleton As DotNetLib.NumberFormatInfoSingleton
End Type

Private this As TNumberFormatInfoSingleton

Private Sub Class_Initialize()
   Set this.NumberFormatInfoSingleton = New DotNetLib.NumberFormatInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.NumberFormatInfoSingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new writable instance of the NumberFormatInfo class that is culture-independent (invariant).")
Public Function Create() As DotNetLib.NumberFormatInfo
Attribute Create.VB_Description = "Initializes a new writable instance of the NumberFormatInfo class that is culture-independent (invariant)."
   Set Create = this.NumberFormatInfoSingleton.Create()
End Function

'@Static
'@Description("Gets a read-only NumberFormatInfo that formats values based on the current culture.")
Public Property Get CurrentInfo() As DotNetLib.NumberFormatInfo
Attribute CurrentInfo.VB_Description = "Gets a read-only NumberFormatInfo that formats values based on the current culture."
   Set CurrentInfo = this.NumberFormatInfoSingleton.CurrentInfo()
End Property

'@Static
'@Description("Gets a read-only NumberFormatInfo object that is culture-independent (invariant).")
Public Property Get InvariantInfo() As DotNetLib.NumberFormatInfo
Attribute InvariantInfo.VB_Description = "Gets a read-only NumberFormatInfo object that is culture-independent (invariant)."
   Set InvariantInfo = this.NumberFormatInfoSingleton.InvariantInfo()
End Property

'@Static
'@Description("Gets the NumberFormatInfo associated with the specified IFormatProvider.")
Public Function GetInstance(ByVal formatProvider As IFormatProvider) As DotNetLib.NumberFormatInfo
Attribute GetInstance.VB_Description = "Gets the NumberFormatInfo associated with the specified IFormatProvider."
   Set GetInstance = this.NumberFormatInfoSingleton.GetInstance(formatProvider)
End Function

'@Static
'@Description("Returns a read-only NumberFormatInfo wrapper.")
Public Function ReadOnly(ByVal nfi As DotNetLib.NumberFormatInfo) As DotNetLib.NumberFormatInfo
Attribute ReadOnly.VB_Description = "Returns a read-only NumberFormatInfo wrapper."
   Set ReadOnly = this.NumberFormatInfoSingleton.ReadOnly(nfi)
End Function
