VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeOffsetStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'Rubberduck annotations
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@Folder("VBADotNetLib.System")
'@PredeclaredId

'@Static

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 11, 2023
'@LastModified July 11, 2023

Option Explicit

Private Type TDateTimeOffsetStatic
    MaxValue As DateTimeOffset
    MinValue As DateTimeOffset
End Type

Private this As TDateTimeOffsetStatic

Private Sub Class_Initialize()
    With New DateTimeOffset
        Set .ComObject = DateTimeOffset.ComObject.MaxValue
        Set this.MaxValue = .Self
    End With
    
    With New DateTimeOffset
        Set .ComObject = DateTimeOffset.ComObject.MinValue
        Set this.MinValue = .Self
    End With
End Sub

'@Static
'@Description("Represents the greatest possible value of DateTimeOffset. This field is read-only.")
Public Property Get MaxValue() As DateTimeOffset
Attribute MaxValue.VB_Description = "Represents the greatest possible value of DateTimeOffset. This field is read-only."
    Set MaxValue = this.MaxValue
End Property

'@Static
'@Description("Represents the earliest possible DateTimeOffset value. This field is read-only.")
Public Property Get MinValue() As DateTimeOffset
Attribute MinValue.VB_Description = "Represents the earliest possible DateTimeOffset value. This field is read-only."
    Set MinValue = this.MinValue
End Property
