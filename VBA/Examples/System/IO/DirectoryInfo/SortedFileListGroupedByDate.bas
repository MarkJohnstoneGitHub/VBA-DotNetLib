Attribute VB_Name = "SortedFileListGroupedByDate"
'@Folder "Examples.System.IO.DirectoryInfo"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 22, 2023
'@LastModified December 24, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.directoryinfo?view=netframework-4.8.1

' https://stackoverflow.com/questions/77694253/what-is-the-fastest-way-to-loop-through-a-folder-with-20-000-files-excel-vba

Option Explicit

' Displays a list of files grouped and sorted according to last write date
' for a provided path, search pattern, search option and filtered for a provided
' start date i.e. the number of days before the current date.
' For each date files and sorted alphabetically by file name.
Public Sub SortedFileListGroupedByDay()
    On Error GoTo ErrorHandler
    
    'inputs for path, search pattern, search option and filter for last write time number of days from today required
    Dim inputPath As String
    inputPath = "C:\VBA\Output"
    Dim inputSearchPattern As String
    inputSearchPattern = "*.cls"
    Dim inputSearchOption As mscorlib.SearchOption
    inputSearchOption = mscorlib.SearchOption.SearchOption_AllDirectories
    Dim inputNumberOfDays As Long
    inputNumberOfDays = 60

    Dim endDate As DotNetLib.DateTime
    Set endDate = DateTime.Today
    Dim startDate As DotNetLib.DateTime
    Set startDate = endDate.AddDays(-inputNumberOfDays)
    
    If Directory.Exists(inputPath) Then
        Dim fileInfos() As DotNetLib.FileSystemInfo
        fileInfos = GetFileSytemInfos(inputPath, inputSearchPattern, inputSearchOption)
    Else
        Err.Raise DirectoryNotFoundException, Description:=VBString.Format("{0} is not a valid file or directory.", inputPath)
    End If
    
    Dim sortedFileList As DotNetLib.SortedList
    Set sortedFileList = GetFilteredSortedFileListGroupedByDay(fileInfos, startDate)
    If sortedFileList.Count = 0 Then
        Debug.Print "No files found."
    Else
        Call DisplayFiles(sortedFileList)
    End If
    Exit Sub
ErrorHandler:
    Debug.Print Err.Description
End Sub

' Obtains an array of FileSystemInfo for a provided path, search pattern and search option
Private Function GetFileSytemInfos(ByVal pPath As String, ByVal pSearchPattern As String, ByVal pSearchOption As mscorlib.SearchOption) As DotNetLib.FileSystemInfo()
    Dim pvtDir As DotNetLib.DirectoryInfo
    Set pvtDir = DirectoryInfo.Create(pPath)
    Dim fileInfos As Variant
    fileInfos = pvtDir.GetFileSystemInfos(pSearchPattern, pSearchOption)
    GetFileSytemInfos = fileInfos
End Function

' Returns a sorted list of file FileSystemInfo by date and grouped according to
' date of last write time and filtered for a provided start date
Private Function GetFilteredSortedFileListGroupedByDay(ByRef fileInfos() As DotNetLib.FileSystemInfo, ByVal startDate As DotNetLib.DateTime) As DotNetLib.SortedList
    Dim pvtIndex As Long
    Dim pvtOutput As DotNetLib.SortedList
    Set pvtOutput = SortedList.Create()
    For pvtIndex = 0 To UBound(fileInfos)
        If fileInfos(pvtIndex).lastWriteTime.Ticks >= startDate.Ticks Then
            Dim daySortedList As DotNetLib.SortedList
            If pvtOutput.ContainsKey(fileInfos(pvtIndex).lastWriteTime.Date) Then
                Set daySortedList = pvtOutput.Item(fileInfos(pvtIndex).lastWriteTime.Date)
            Else
                Set daySortedList = SortedList.Create()
                Call pvtOutput.Add(fileInfos(pvtIndex).lastWriteTime.Date, daySortedList)
            End If
            'sorted key is last write time and full name to avoid potiential issue of duplicate key due to last write time
            'Could create a custom sort using an IComparer for the daily sorted file list
            Call daySortedList.Add(fileInfos(pvtIndex).name & "," & fileInfos(pvtIndex).lastWriteTime.Date.Ticks, fileInfos(pvtIndex))
        End If
    Next
    Set GetFilteredSortedFileListGroupedByDay = pvtOutput
End Function

Private Sub DisplayFiles(ByVal pList As DotNetLib.SortedList)
    Dim pvtFormat As String
    pvtFormat = "{0}, Last Modified: {1}"

    Dim i As Long
    For i = pList.Count - 1 To 0 Step -1 'Transverse in reverse order from end date to start date
        'day list
        Dim daySortedFileList As DotNetLib.SortedList
        Set daySortedFileList = pList.GetByIndex(i)
        Debug.Print VBString.Format("Files last modified on {0:d}", pList.GetKey(i))
        
        Dim j As Long
        For j = 0 To daySortedFileList.Count - 1
            Dim pvtfileInfo As DotNetLib.FileSystemInfo
            Set pvtfileInfo = daySortedFileList.GetByIndex(j)
            Debug.Print VBString.Format(pvtFormat, pvtfileInfo.name, pvtfileInfo.lastWriteTime)
        Next j
        Debug.Print
    Next i
End Sub
