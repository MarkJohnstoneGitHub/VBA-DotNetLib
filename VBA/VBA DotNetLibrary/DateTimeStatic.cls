VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("DotNet.System")
'@PredeclaredId

Option Explicit

Private Type TDateTimeStatic
    MaxValue As DateTime
    MinValue As DateTime
End Type

Private this As TDateTimeStatic

Private Sub Class_Initialize()
    Set this.MaxValue = New DateTime
    Set this.MaxValue.ComObject = DateTime.ComObject.MaxValue
    
    Set this.MinValue = New DateTime
    Set this.MinValue.ComObject = DateTime.ComObject.MinValue
End Sub

Public Property Get MaxValue() As DateTime
    Set MaxValue = this.MaxValue
End Property

Public Property Get MinValue() As DateTime
    Set MinValue = this.MinValue
End Property

