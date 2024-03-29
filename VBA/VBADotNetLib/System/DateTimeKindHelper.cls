VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeKindHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "DateTimeKind helper class to convert a DateTimeKind to string value."
'Rubberduck annotations
'@ModuleDescription("DateTimeKind helper class to convert a DateTimeKind to string value.")
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 31, 2023

'@ReferenceAddin DotNetLib.tlb

'@Notes:
' https://www.codevba.com/help/enum_builder.htm
' https://stackoverflow.com/questions/41578348/is-there-a-way-to-get-the-enums-in-vba
' https://stackoverflow.com/a/41579986/10759363
' Possible could allow for various culture/language by obtaining a two dimensional array
' containing translations or method convert to another language?
' Would require Class_Initialize() to populate a language array for DateTimeKind from a resource file?

Option Explicit

'@DefaultMember
'@Description("Converts the value of a DateTimeKind to string value.")
Public Property Get ToString(ByVal value As DotNetLib.DateTimeKind) As String
Attribute ToString.VB_Description = "Converts the value of a DateTimeKind to string value."
Attribute ToString.VB_UserMemId = 0
   Select Case value
      Case DateTimeKind_Local: ToString = "Local"
      Case DateTimeKind_Utc: ToString = "UTC"
      Case DateTimeKind_Unspecified: ToString = "Unspecified"
   End Select
End Property
