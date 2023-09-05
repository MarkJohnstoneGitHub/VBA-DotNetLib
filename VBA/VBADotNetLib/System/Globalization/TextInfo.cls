VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "TextInfo factory methods and static members."
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("VBADotNetLib.System.Globalization")
'@PredeclaredId
'@ModuleDescription("TextInfo factory methods and static members.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 05, 2023
'@LastModified September 05, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo?view=netframework-4.8.1

Option Explicit

Private Type TTextInfo
   TextInfoSingleton As DotNetLib.TextInfoSingleton
End Type

Private this As TTextInfo

Private Sub Class_Initialize()
   Set this.TextInfoSingleton = New DotNetLib.TextInfoSingleton
End Sub

Private Sub Class_Terminate()
   Set this.TextInfoSingleton = Nothing
End Sub

'@Static
'@Description("Returns a read-only version of the specified TextInfo object.")
Public Function ReadOnly(ByVal pTextInfo As DotNetLib.TextInfo) As DotNetLib.TextInfo
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified TextInfo object."
    Set ReadOnly = this.TextInfoSingleton.ReadOnly(pTextInfo)
End Function
