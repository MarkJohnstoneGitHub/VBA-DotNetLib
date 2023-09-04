Attribute VB_Name = "DTFIGetAllDateTimePatternsEg2"
'@Folder("Examples.System.Globalization.DateTimeFormatInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 4, 2023
'@LastModified September 4, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.datetimeformatinfo.getalldatetimepatterns?view=netframework-4.8.1#system-globalization-datetimeformatinfo-getalldatetimepatterns(system-char)

Option Explicit

Public Sub DateTimeFormatInfoGetAllDateTimePatterns2()
    Dim myDtfi As DotNetLib.DateTimeFormatInfo
    Set myDtfi = DateTimeFormatInfo.Create()
    
    ' Gets and prints all the patterns.
    Dim myPatternsArray() As String
    myPatternsArray = myDtfi.GetAllDateTimePatterns()
    Debug.Print "ALL the patterns:"
    PrintIndexAndValues myPatternsArray
    
    ' Gets and prints the pattern(s) associated with some of the format characters.
    myPatternsArray = myDtfi.GetAllDateTimePatterns("d")
    Debug.Print "The patterns for 'd':"
    PrintIndexAndValues myPatternsArray
    
    myPatternsArray = myDtfi.GetAllDateTimePatterns("D")
    Debug.Print "The patterns for 'D':"
    PrintIndexAndValues myPatternsArray
    
    myPatternsArray = myDtfi.GetAllDateTimePatterns("f")
    Debug.Print "The patterns for 'f':"
    PrintIndexAndValues myPatternsArray
    
    myPatternsArray = myDtfi.GetAllDateTimePatterns("F")
    Debug.Print "The patterns for 'F':"
    PrintIndexAndValues myPatternsArray
    
    myPatternsArray = myDtfi.GetAllDateTimePatterns("r")
    Debug.Print "The patterns for 'r':"
    PrintIndexAndValues myPatternsArray
    
    myPatternsArray = myDtfi.GetAllDateTimePatterns("R")
    Debug.Print "The patterns for 'R':"
    PrintIndexAndValues myPatternsArray
    
End Sub

Private Sub PrintIndexAndValues(ByRef myArray() As String)
    Dim i As Long
    Dim s As Variant
    For Each s In myArray
        Debug.Print VBA.vbTab; "[" & i & "]"; VBA.vbTab; s
        i = i + 1
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'ALL the patterns:
'    [0]    MM / dd / yyyy
'    [1]    yyyy - MM - dd
'    [2]    dddd, dd MMMM yyyy
'    [3]    dddd, dd MMMM yyyy HH:mm
'    [4]    dddd, dd MMMM yyyy hh:mm tt
'    [5]    dddd, dd MMMM yyyy H:mm
'    [6]    dddd, dd MMMM yyyy h:mm tt
'    [7]    dddd, dd MMMM yyyy HH:mm:ss
'    [8]    MM/dd/yyyy HH:mm
'    [9]    MM/dd/yyyy hh:mm tt
'    [10]   MM/dd/yyyy H:mm
'    [11]   MM/dd/yyyy h:mm tt
'    [12]   yyyy-MM-dd HH:mm
'    [13]   yyyy-MM-dd hh:mm tt
'    [14]   yyyy-MM-dd H:mm
'    [15]   yyyy-MM-dd h:mm tt
'    [16]   MM/dd/yyyy HH:mm:ss
'    [17]   yyyy-MM-dd HH:mm:ss
'    [18]   MMMM dd
'    [19]   MMMM dd
'    [20]   yyyy   '-'MM'-'dd'T'HH':'mm':'ss.fffffffK
'    [21]   yyyy   '-'MM'-'dd'T'HH':'mm':'ss.fffffffK
'    [22]   ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'    [23]   ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'    [24]   yyyy   '-'MM'-'dd'T'HH':'mm':'ss
'    [25]   HH: MM
'    [26]   HH: MM tt
'    [27]   h: MM
'    [28]   h: MM tt
'    [29]   HH: MM: ss
'    [30]   yyyy   '-'MM'-'dd HH':'mm':'ss'Z'
'    [31]   dddd, dd MMMM yyyy HH:mm:ss
'    [32]   yyyy MMMM
'    [33]   yyyy MMMM
'
'The patterns for 'd':
'    [0]    MM / dd / yyyy
'    [1]    yyyy - MM - dd
'
'The patterns for 'D':
'    [0]    dddd, dd MMMM yyyy
'
'The patterns for 'f':
'    [0]    dddd, dd MMMM yyyy HH:mm
'    [1]    dddd, dd MMMM yyyy hh:mm tt
'    [2]    dddd, dd MMMM yyyy H:mm
'    [3]    dddd, dd MMMM yyyy h:mm tt
'
'The patterns for 'F':
'    [0]    dddd, dd MMMM yyyy HH:mm:ss
'
'The patterns for 'r':
'    [0]    ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'
'The patterns for 'R':
'    [0]    ddd, dd MMM yyyy HH':'mm':'ss 'GMT'
'
'*/
