Attribute VB_Name = "TimeSpanTryParseExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.tryparse?view=netframework-4.8.1#system-timespan-tryparse(system-string-system-timespan@)

Option Explicit

''
' The following example uses the TryParse method to create TimeSpan objects from
' valid TimeSpan strings and to indicate when the parse operation has failed
' because the time span string is invalid.
''
Public Sub TimeSpanTryParse()
    Debug.Print VBString.Format("{0,20}   {1,21}", _
        "String to Parse", "TimeSpan")
    Debug.Print VBString.Format("{0,20}   {1,21}", _
        "---------------", "---------------------")
        
    ParseTimeSpan "0"
    ParseTimeSpan "14"
    ParseTimeSpan "1:2:3"
    ParseTimeSpan "0:0:0.250"
    ParseTimeSpan "10.20:30:40.50"
    ParseTimeSpan "99.23:59:59.9999999"
    ParseTimeSpan "0023:0059:0059.0099"
    ParseTimeSpan "23:0:0"
    ParseTimeSpan "24:0:0"
    ParseTimeSpan "0:59:0"
    ParseTimeSpan "0:60:0"
    ParseTimeSpan "0:0:59"
    ParseTimeSpan "0:0:60"
    ParseTimeSpan "10:"
    ParseTimeSpan "10:0"
    ParseTimeSpan ":10"
    ParseTimeSpan "0:10"
    ParseTimeSpan "10:20:"
    ParseTimeSpan "10:20:0"
    ParseTimeSpan ".123"
    ParseTimeSpan "0.12:00"
    ParseTimeSpan "10."
    ParseTimeSpan "10.12"
    ParseTimeSpan "10.12:00"
End Sub

Private Sub ParseTimeSpan(ByVal intervalStr As String)
    ' Write the first part of the output line.
    Debug.Print VBString.Format("{0,20}   ", intervalStr);

    ' Parse the parameter, and then convert it back to a string.
    Dim intervalVal As DotNetLib.TimeSpan
    
    If (TimeSpan.TryParse(intervalStr, intervalVal)) Then
        Dim intervalToStr As String
        intervalToStr = intervalVal.ToString()
        
        ' Pad the end of the TimeSpan string with spaces if it
        ' does not contain milliseconds.
        Dim pIndex As Long
        pIndex = InStr(intervalToStr, ":")
        pIndex = InStr(pIndex, intervalToStr, ".")
        If (pIndex = 0) Then
            intervalToStr = intervalToStr & "        "
        End If
    
        Debug.Print VBString.Format("{0,21}", intervalToStr)
    Else  ' Handle failure of TryParse method.
        Debug.Print "Parse operation failed."
    End If
End Sub

' Output:
'   String to Parse                          TimeSpan
'   ---------------             ---------------------
'   0                              00:00:00
'   14                          14.00:00:00
'   1:2:3                          01:02:03
'   0:0:0.250                      00:00:00.2500000
'   10.20:30:40.50              10.20:30:40.5000000
'   99.23:59:59.9999999         99.23:59:59.9999999
'   0023:0059:0059.0099            23:59:59.0099000
'   23:0:0                         23:00:00
'   24:0:0                      24.00:00:00
'   0:59:0                         00:59:00
'   0:60:0                      Parse operation failed.
'   0:0:59                         00:00:59
'   0:0:60                      Parse operation failed.
'   10:                         Parse operation failed.
'   10:0                           10:00:00
'   :10                         Parse operation failed.
'   0:10                           00:10:00
'   10:20:                      Parse operation failed.
'   10:20:0                        10:20:00
'   .123                        Parse operation failed.
'   0.12:00                        12:00:00
'   10.                         Parse operation failed.
'   10.12                       Parse operation failed.
'   10.12:00                    10.12:00:00

