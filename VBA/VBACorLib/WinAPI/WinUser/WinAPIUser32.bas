Attribute VB_Name = "WinAPIUser32"
'@Folder("VBACorLib.WinAPI.WinUser")

Option Explicit

' https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-messageboxw
' http://blog.nkadesign.com/2013/10/01/vba-unicode-strings-and-the-windows-api/
#If Not Mac And VBA7 Then
    Public Declare PtrSafe Function MessageBoxW Lib "user32" _
        (ByVal hwnd As LongPtr, _
         ByVal lpText As LongPtr, _
         ByVal lpCaption As LongPtr, _
         ByVal wType As Long) As Long
#ElseIf Not Mac Then
    Public Declare Function MessageBoxW Lib "user32" Alias "MessageBoxW" _
        (ByVal hwnd As Long, _
         ByVal lpText As Long, _
         ByVal lpCaption As Long, _
         ByVal wType As Long) As Long
#End If
