Attribute VB_Name = "DateTimeExceptions"
'@Folder("VBADotNetLib.Examples.DateTime")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12, 2023
'@LastModified July 12, 2023

' Notes
' https://www.automateexcel.com/vba/error-handling/

Option Explicit

Public Sub DateTimeCreateFromTicksException()
   Dim ticksErr As LongLong
   ticksErr = DateTime.MaxValue.Ticks + 1
   On Error Resume Next    'Skip error and continue running
   Dim dateValue As DateTime
   Set dateValue = DateTime.CreateFromTicks(ticksErr)
   If Err.Number = COMHResult.ArgumentOutOfRangeException Then
      Debug.Print "ArgumentOutOfRangeException " & "0x" & Hex$(Err.Number) & " " & Err.Description
   Else
      Debug.Print "0x" & Hex$(Err.Number) & " " & Err.Description
   End If
   On Error GoTo 0 'Stop code and display error
' Output:
' ArgumentOutOfRangeException 0x80131502 Ticks must be between DateTime.MinValue.Ticks and DateTime.MaxValue.Ticks.
' Parameter Name: Ticks
End Sub


Public Sub DateTimeCreateFromDateTimeKindException()
   On Error Resume Next    'Skip error and continue running
   Dim dateValue As DateTime
   Set dateValue = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, 5)
   If Err.Number = COMHResult.ArgumentException Then
      Debug.Print "ArgumentException " & "0x" & Hex$(Err.Number) & " " & Err.Description
   Else
      Debug.Print "0x" & Hex$(Err.Number) & " " & Err.Description
   End If
   On Error GoTo 0   'Stop code and display error

' Output:
' ArgumentException 0x80070057 Invalid DateTimeKind value.
' Parameter Name: kind
End Sub
