Attribute VB_Name = "DateTimeExceptions"
'@IgnoreModule VariableNotUsed
'@Folder "Examples.System.DateTime.Exceptions"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 12, 2023
'@LastModified August 4, 2023

'@Notes
' https://www.automateexcel.com/vba/error-handling/

Option Explicit

Public Sub DateTimeCreateFromTicksException()
   Dim ticksErr As LongLong
   ticksErr = DateTime.MaxValue.Ticks + 1
   On Error Resume Next    'Skip error and continue running
   Dim dateValue As IDateTime
   Set dateValue = DateTime.CreateFromTicks(ticksErr)
   If Err.number = COMHResult.ArgumentOutOfRangeException Then
      Debug.Print "ArgumentOutOfRangeException " & "0x" & Hex$(Err.number) & " " & Err.Description
   Else
      Debug.Print "0x" & Hex$(Err.number) & " " & Err.Description
   End If
   On Error GoTo 0 'Stop code and display error
End Sub

' Output:
' ArgumentOutOfRangeException 0x80131502 Ticks must be between DateTime.MinValue.Ticks and DateTime.MaxValue.Ticks.
' Parameter Name: Ticks

Public Sub DateTimeCreateFromDateTimeKindException()
   On Error Resume Next    'Skip error and continue running
   Dim dateValue As IDateTime
   Set dateValue = DateTime.CreateFromDateTimeKind(2010, 8, 18, 16, 32, 0, 5)
   If Err.number = COMHResult.ArgumentException Then
      Debug.Print "ArgumentException " & "0x" & Hex$(Err.number) & " " & Err.Description
   Else
      Debug.Print "0x" & Hex$(Err.number) & " " & Err.Description
   End If
   On Error GoTo 0   'Stop code and display error
End Sub

' Output:
' ArgumentException 0x80070057 Invalid DateTimeKind value.
' Parameter Name: kind
