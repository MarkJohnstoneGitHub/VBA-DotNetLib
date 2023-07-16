Attribute VB_Name = "ErrorHandling"
'@Folder("VBACorLib.ErrorHandling")

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 13, 2023
'@LastModified July 17, 2023

'@TODO Work in progress

'@Notes
' https://learn.microsoft.com/en-us/dotnet/standard/exceptions/how-to-create-localized-exception-messages

Option Explicit

Public Function Try() As Boolean
   If Err.Number = 0 Then
      Try = True
   End If
End Function

'@TODO Pass in an error object i.e. Exception eg. FormatError
'Public Function Catch(Optional ByVal errorInfo As IException) As Boolean
Public Function Catch(Optional ByVal error As COMHResult) As Boolean
   If Err.Number = error Then
      Catch = True
   Else
      If error = 0 Then 'i.e. optional
         If Err.Number <> 0 Then
            Catch = True
         End If
      End If
   End If
End Function



