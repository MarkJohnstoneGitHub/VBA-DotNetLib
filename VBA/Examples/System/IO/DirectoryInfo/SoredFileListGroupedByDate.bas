Attribute VB_Name = "SoredFileListGroupedByDate"
'@Folder("Examples.System.IO.DirectoryInfo")
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 22, 2023
'@LastModified December 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

' https://stackoverflow.com/questions/77694253/what-is-the-fastest-way-to-loop-through-a-folder-with-20-000-files-excel-vba

'@TODO Fix SortedList enumeration to enable use For Each to obtain a Enumerator of IEnumVariant
' i.e update the SortList.GetEnumerator to return IEnumVariant

Option Explicit

' Displays a list of files grouped and sorted according to last write date
' for a provided path, search pattern, search option and filtered for a provided
' start date i.e. the number of days before the current date.
Public Sub SortedFileListGroupedByDay()
    On Error GoTo ErrorHandler
    
    'inputs for path, search pattern, search option and filter for last write time number of days from today required
    Dim pvtFolderPath As String
    pvtFolderPath = "C:\VBA\Output"
    Dim pvtSearchPattern As String
    pvtSearchPattern = "*.cls"
    Dim pvtSearchOption As mscorlib.SearchOption
    pvtSearchOption = mscorlib.SearchOption.SearchOption_AllDirectories
    Dim pvtNumberOfDays As Long
    pvtNumberOfDays = 60

    
    Dim pvtEndDate As DotNetLib.DateTime
    Set pvtEndDate = DateTime.Today
    Dim pvtStartDate As DotNetLib.DateTime
    Set pvtStartDate = pvtEndDate.AddDays(-pvtNumberOfDays)
    
    If Directory.Exists(pvtFolderPath) Then
        Dim fileInfos() As DotNetLib.FileSystemInfo
        fileInfos = GetFileSytemInfos(pvtFolderPath, pvtSearchPattern, pvtSearchOption)
    Else
        Err.Raise DirectoryNotFoundException, Description:=VBAString.Format("{0} is not a valid file or directory.", pvtFolderPath)
    End If
    
    Dim pvtFilesSortedList As DotNetLib.SortedList
    Set pvtFilesSortedList = GetFileredSortedFileListByDay(fileInfos, pvtStartDate)
    
    If pvtFilesSortedList.Count = 0 Then
        Debug.Print "No files found."
    Else
        Call DisplayFiles(pvtFilesSortedList)
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

Private Function GetFileSytemInfos(ByVal pPath As String, ByVal pSearchPattern As String, ByVal pSearchOption As mscorlib.SearchOption) As DotNetLib.FileSystemInfo()
    Dim pvtDir As DotNetLib.DirectoryInfo
    Set pvtDir = DirectoryInfo.Create(pPath)
    Dim infos() As DotNetLib.FileSystemInfo
    infos = pvtDir.GetFileSystemInfos(pSearchPattern, pSearchOption)
    GetFileSytemInfos = infos
End Function

'Returns a list of file FileSystemInfo grouped and sorted according to last write date filtered for a provided start date
Private Function GetFileredSortedFileListByDay(ByRef fileInfos() As DotNetLib.FileSystemInfo, ByVal startDate As DotNetLib.DateTime) As DotNetLib.SortedList
    Dim pvtIndex As Long
    Dim pvtOutput As DotNetLib.SortedList
    Set pvtOutput = SortedList.Create()
    For pvtIndex = 0 To UBound(fileInfos)
        If fileInfos(pvtIndex).lastWriteTime.Date.Ticks >= startDate.Date.Ticks Then
            Dim daySortedList As DotNetLib.SortedList
            'Require to also filter by file name
            
            If Not pvtOutput.ContainsKey(fileInfos(pvtIndex).lastWriteTime.Date) Then
                Set daySortedList = SortedList.Create()
                Call pvtOutput.Add(fileInfos(pvtIndex).lastWriteTime.Date, daySortedList)
            Else
                Set daySortedList = pvtOutput.Item(fileInfos(pvtIndex).lastWriteTime.Date)
            End If
            'sorted key is last write time and full name to avoid potiential issue of duplicate last write time
            Call daySortedList.Add(CStr(fileInfos(pvtIndex).lastWriteTime.Date & "," & fileInfos(pvtIndex).FullName), fileInfos(pvtIndex))
        End If
    Next
    Set GetFileredSortedFileListByDay = pvtOutput
End Function

'@TODO Fix SortedList enumeration to enable use For Each requires obtaining a Enumerator of IEnumVariant
Private Sub DisplayFiles(ByVal pList As DotNetLib.SortedList)
    Dim pvtFormatString As String
    pvtFormatString = "{0}, Last Modified: {1}"
    Dim pvtValueList As mscorlib.IList
    Set pvtValueList = pList.GetValueList()
    
    Dim i As Long
    For i = pList.Count - 1 To 0 Step -1 'Transverse in reverse order from end date to start date
        'day list
        Dim daySortedFileList As DotNetLib.SortedList
        Set daySortedFileList = pvtValueList(i)
        Debug.Print VBAString.Format("Files last modified on {0:d}", pList.GetKey(i))
        
        Dim j As Long
        For j = 0 To daySortedFileList.Count - 1
            Dim pvtFileSystemInfo As DotNetLib.FileSystemInfo
            Set pvtFileSystemInfo = daySortedFileList.GetByIndex(j)
            Debug.Print VBAString.Format(pvtFormatString, pvtFileSystemInfo.name, pvtFileSystemInfo.lastWriteTime)
        Next j
        Debug.Print
    Next i
End Sub
