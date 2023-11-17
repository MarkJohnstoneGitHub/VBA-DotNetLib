Attribute VB_Name = "DirectoryGetCreationTimeEg"
'@Folder("Examples.System.IO.Directory.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 17, 2023
'@LastModified November 17, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directory.getcreationtime?view=netframework-4.8.1

Option Explicit

''
' The following example gets the creation time of the specified directory.
''
Public Sub DirectoryGetCreationTime()
    On Error GoTo ErrorHandler
    ' Get the creation time of a well-known directory.
    Dim dt As DotNetLib.DateTime
    Set dt = Directory.GetCreationTime(Environment.CurrentDirectory)
    Debug.Print Environment.CurrentDirectory
    ' Give feedback to the user.
    If (DateTime.Now.Subtract2(dt).TotalDays > 364) Then
        Debug.Print "This directory is over a year old."
    ElseIf (DateTime.Now.Subtract2(dt).TotalDays > 30) Then
        Debug.Print "This directory is over a month old."
    ElseIf (DateTime.Now.Subtract2(dt).TotalDays <= 1) Then
        Debug.Print "This directory is less than a day old."
    Else
        Debug.Print VBAString.Format("This directory was created on {0}", dt)
    End If
Exit Sub

ErrorHandler:
    Debug.Print VBAString.Format("The process failed: {0}", Err.Description)
End Sub
