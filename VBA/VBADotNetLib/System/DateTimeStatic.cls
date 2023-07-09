VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rubberduck annotations
'@IgnoreModule UnrecognizedAnnotation
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Static

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified July 09, 2023

Option Explicit

Private Type TDateTimeStatic
    MaxValue As DateTime
    MinValue As DateTime
End Type

Private this As TDateTimeStatic

Private Sub Class_Initialize()
    With New DateTime
        Set .ComObject = DateTime.ComObject.MaxValue
        Set this.MaxValue = .Self
    End With
    
    With New DateTime
        Set .ComObject = DateTime.ComObject.MinValue
        Set this.MinValue = .Self
    End With
End Sub

'@Static
'@Description("Represents the largest possible value of DateTime. This field is read-only.")
Public Property Get MaxValue() As DateTime
Attribute MaxValue.VB_Description = "Represents the largest possible value of DateTime. This field is read-only."
    Set MaxValue = this.MaxValue
End Property

'@Static
'@Description("Represents the smallest possible value of DateTime. This field is read-only.")
Public Property Get MinValue() As DateTime
Attribute MinValue.VB_Description = "Represents the smallest possible value of DateTime. This field is read-only."
    Set MinValue = this.MinValue
End Property
