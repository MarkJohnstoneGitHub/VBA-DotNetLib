Attribute VB_Name = "StringFormatting"
'@Folder("VBACorLib.Formatting")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 August 1, 2023
'@LastModified August 4, 2023


Option Explicit

'@TODO StringFormatting Align

Public Enum Justify
    Justify_Right = 1
    Justify_Left = 2
End Enum

'@Description("String alignment left or right justify for a provided width")
Public Function Align(ByVal text As String, ByVal width As Long, Optional ByVal alignment As Justify = Justify.Justify_Left) As String
Attribute Align.VB_Description = "String alignment left or right justify for a provided width"
    
' Test if width is greater then text string length? width>=1? Allow truncation? Throw errors or just return text back?
'    If width < 1 Then
'       Err.Raise
'    End If
'
    Select Case alignment
        Case Justify.Justify_Left
            Align = Left$(Space(width) & text, width)
        Case Justify.Justify_Right
            Align = Right$(Space(width) & text, width)
        Case Else
            Align = text    'Return error invalid parameter or return string unformatted?
    End Select
End Function
