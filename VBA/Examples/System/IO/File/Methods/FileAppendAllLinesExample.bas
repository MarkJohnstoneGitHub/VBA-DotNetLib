Attribute VB_Name = "FileAppendAllLinesExample"
'@Folder("Examples.System.IO.File.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 November 21, 2023
'@LastModified November 21, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.io.file.appendalllines?view=netframework-4.8.1

Option Explicit

Private Const dataPath As String = "c:\temp\timestamps.txt"

''
' The following example writes selected lines from a sample data file to a file,
' and then appends more lines.
' The directory named temp on drive C must exist for the example to complete successfully.
''
Public Sub FileAppendAllLinesExample()
    Call CreateSampleFile
    Call file.WriteAllLines3("C:\temp\selectedDays.txt", JulyWeekends)
    Call file.AppendAllLines("C:\temp\selectedDays.txt", MarchMondays)
End Sub

Private Sub CreateSampleFile()
    Dim TimeStamp As DotNetLib.DateTime
    Set TimeStamp = DateTime.CreateFromDate(1700, 1, 1)

    Dim sw As DotNetLib.StreamWriter
    Set sw = StreamWriter.Create(dataPath)
    Dim i As Long
    For i = 0 To 499
        Dim TS1 As DotNetLib.DateTime
        Set TS1 = TimeStamp.AddYears(i)
        Dim TS2 As DotNetLib.DateTime
        Set TS2 = TS1.AddMonths(i)
        Dim TS3 As DotNetLib.DateTime
        Set TS3 = TS2.AddDays(i)
        Call sw.WriteLine2(TS3.ToLongDateString())
    Next
    Call sw.Dispose
End Sub

Private Function JulyWeekends() As mscorlib.IEnumerable
    Dim output As DotNetLib.ListString
    Set output = ListString.Create()
    
    Dim varLine As Variant
    For Each varLine In file.ReadLines(dataPath)
        Dim line As DotNetLib.String
        Set line = Strings.Copy(varLine)
        If (line.StartsWith_4("Saturday") Or line.StartsWith_4("Sunday")) And line.Contains_2("July") Then
            Call output.Add(line.ToString())
        End If
    Next
    Set JulyWeekends = output.GetIEnumerable
End Function

Private Function MarchMondays() As mscorlib.IEnumerable
    Dim output As DotNetLib.ListString
    Set output = ListString.Create()
        
    Dim varLine As Variant
    For Each varLine In file.ReadLines(dataPath)
        Dim line As DotNetLib.String
        Set line = Strings.Copy(varLine)
        If (line.StartsWith_4("Monday") And line.Contains_2("March")) Then
            Call output.Add(line.ToString())
        End If
    Next
    Set MarchMondays = output.GetIEnumerable
End Function
