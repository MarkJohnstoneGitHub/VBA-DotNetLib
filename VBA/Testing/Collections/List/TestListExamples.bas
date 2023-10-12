Attribute VB_Name = "TestListExamples"
'@Folder("Testing.Collections.List")
Option Explicit

Private Sub TestList()
    Dim pvtDates As DotNetLib.List
    Dim listType As DotNetLib.DateTime
    Set listType = New DotNetLib.DateTime 'Require to set a "default" value so can determine type
    Set pvtDates = List.Create2(listType)
    pvtDates.Add DateTime.CreateFromDate(2023, 1, 1)
    pvtDates.Add DateTime.CreateFromDate(2022, 2, 1)
End Sub

Private Sub TestListExample2()
    Dim longList As DotNetLib.List
    Dim listType As Long
    Set longList = List.Create2(listType)
    'Note must convert to long or invalid type
    longList.Add CLng(10)
    longList.Add CLng(55)
    longList.Add CLng(5)
    longList.Sort 'Error casting??? RunTime Error '-2147467262 (800040002) Unable to cast System.Generic.List'1[System.Int32]' to type 'System.Generic.List'1[System.Object]'.
End Sub

