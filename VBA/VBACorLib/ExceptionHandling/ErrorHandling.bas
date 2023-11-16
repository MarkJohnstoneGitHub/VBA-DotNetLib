Attribute VB_Name = "ErrorHandling"
'@Folder "VBACorLib.ExceptionHandling"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 13, 2023
'@LastModified August 24, 2023

'@TODO Work in progress

'@Notes
' https://learn.microsoft.com/en-us/dotnet/standard/exceptions/how-to-create-localized-exception-messages

Option Explicit

Public Function Try() As Boolean
   If Err.number = 0 Then
      Try = True
   End If
End Function

'Public Function Catch(Optional ByVal errorInfo As IException) As Boolean
'@TODO Pass in an error object i.e. Exception eg. FormatError
'Public Function Catch(Optional ByVal errorInfo As IException) As Boolean
Public Function Catch(Optional ByVal errNumber As Variant) As Boolean
    If IsMissing(errNumber) Then
        If Err.number <> 0 Then
            Catch = True
        End If
    ElseIf Err.number = errNumber Then
        Catch = True
    End If
End Function

