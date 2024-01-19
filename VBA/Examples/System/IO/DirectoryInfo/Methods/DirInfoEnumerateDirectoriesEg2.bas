Attribute VB_Name = "DirInfoEnumerateDirectoriesEg2"
'@Folder "Examples.System.IO.DirectoryInfo.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 25, 2023
'@LastModified December 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo.enumeratedirectories?view=netframework-4.8.1#system-io-directoryinfo-enumeratedirectories(system-string-system-io-searchoption)

Option Explicit

Public Sub DirectoryInfoEnumerateDirectoriesExample2()
    On Error Resume Next

    ' Set a variable to the My Documents path.
    Dim docPath As String
    docPath = Environment.GetFolderPath(SpecialFolder.SpecialFolder_MyDocuments)

    Dim diTop As DotNetLib.DirectoryInfo
    Set diTop = DirectoryInfo.Create(docPath)
    
    If Err.Number = 0 Then
        Dim pvtFormatString As String
        pvtFormatString = Regex.Unescape("{0}\t\t{1:N0}")
        Dim varFileInfo As Variant
        For Each varFileInfo In diTop.EnumerateFiles()
            If Err.Number = 0 Then
                Dim fi As DotNetLib.FileInfo
                Set fi = varFileInfo
                ' Display each file over 10 MB.
                If (fi.Length > 10000000) Then
                    Debug.Print VBString.Format(pvtFormatString, fi.FullName, fi.Length)
                End If
            ElseIf Err.Number = UnauthorizedAccessException Then
                Debug.Print Err.Description
            End If
        Next
    
        Dim varDirInfo As Variant
        For Each varDirInfo In diTop.EnumerateDirectories("*")
            If Err.Number = 0 Then
                Dim di As DotNetLib.DirectoryInfo
                Set di = varDirInfo
                For Each varFileInfo In di.EnumerateFiles("*", SearchOption.SearchOption_AllDirectories)
                    Set fi = varFileInfo
                    If Err.Number = 0 Then
                        Set fi = varFileInfo
                        ' Display each file over 10 MB.
                        If (fi.Length > 10000000) Then
                            Debug.Print VBString.Format(pvtFormatString, fi.FullName, fi.Length)
                        End If
                    ElseIf Err.Number = UnauthorizedAccessException Then
                        Debug.Print "unAuthFile: " & Err.Description
                    End If
                Next
            ElseIf Err.Number = UnauthorizedAccessException Then
                Debug.Print "unAuthSubDir: " & Err.Description
            End If
        Next
    ElseIf Err.Number = DirectoryNotFoundException Then
        Debug.Print Err.Description
    ElseIf Err.Number = UnauthorizedAccessException Then
        Debug.Print "unAuthDir: " & Err.Description
    ElseIf Err.Number = PathTooLongException Then
        Debug.Print Err.Description
    End If
    On Error GoTo 0 'Stop code and display error
End Sub
