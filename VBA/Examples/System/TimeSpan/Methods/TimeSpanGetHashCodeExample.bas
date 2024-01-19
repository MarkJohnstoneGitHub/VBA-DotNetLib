Attribute VB_Name = "TimeSpanGetHashCodeExample"
'@Folder "Examples.System.TimeSpan.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified January 18, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.gethashcode?view=netframework-4.8.1#examples

Option Explicit

''
' The following example generates the hash codes of several TimeSpan objects
' using the GetHashCode method.
''
Public Sub TimeSpanGetHashCode()
    Debug.Print VBString.Unescape( _
        "This example of TimeSpan.GetHashCode( ) generates " + _
        "the following \noutput, which displays " + _
        "the hash codes of representative TimeSpan \n" + _
        "objects in hexadecimal and decimal formats.\n")
    Debug.Print VBString.Format("{0,22}   {1,10}", _
        "TimeSpan        ", "Hash Code")
    Debug.Print VBString.Format("{0,22}   {1,10}", _
        "--------        ", "---------")
   
    DisplayHashCode TimeSpan.CreateFromTicks(0)
    DisplayHashCode TimeSpan.CreateFromTicks(1)
    DisplayHashCode TimeSpan.Create3(0, 0, 0, 0, 1)
    DisplayHashCode TimeSpan.Create(0, 0, 1)
    DisplayHashCode TimeSpan.Create(0, 1, 0)
    DisplayHashCode TimeSpan.Create(1, 0, 0)
    DisplayHashCode TimeSpan.CreateFromTicks(36000000001#)
    DisplayHashCode TimeSpan.Create3(0, 1, 0, 0, 1)
    DisplayHashCode TimeSpan.Create(1, 0, 1)
    DisplayHashCode TimeSpan.Create2(1, 0, 0, 0)
    DisplayHashCode TimeSpan.CreateFromTicks(864000000001#)
    DisplayHashCode TimeSpan.Create3(1, 0, 0, 0, 1)
    DisplayHashCode TimeSpan.Create2(1, 0, 0, 1)
    DisplayHashCode TimeSpan.Create2(100, 0, 0, 0)
    DisplayHashCode TimeSpan.Create3(100, 0, 0, 0, 1)
    DisplayHashCode TimeSpan.Create2(100, 0, 0, 1)
End Sub

Private Sub DisplayHashCode(ByVal interval As DotNetLib.TimeSpan)
    ' Create a hash code and a string representation of
    ' the TimeSpan parameter.
    Dim timeInterval As String
    timeInterval = interval.ToString()
    Dim hashCode As Long
    hashCode = interval.GetHashCode()
   
    ' Pad the end of the TimeSpan string with spaces if it
    ' does not contain milliseconds.
    Dim pIndex As Long
    pIndex = InStr(timeInterval, ":")
    pIndex = InStr(pIndex, timeInterval, ".")
    If (pIndex = 0) Then
        timeInterval = timeInterval & "        "
    End If
    
    Debug.Print VBString.Format("{0,22}   0x{1:X8}, {1}", _
        timeInterval, hashCode)
End Sub

'/*
'This example of TimeSpan.GetHashCode( ) generates the following
'output, which displays the hash codes of representative TimeSpan
'objects in hexadecimal and decimal formats.
'
'      TimeSpan            Hash Code
'      --------            ---------
'      00:00:00           0x00000000, 0
'      00:00:00.0000001   0x00000001, 1
'      00:00:00.0010000   0x00002710, 10000
'      00:00:01           0x00989680, 10000000
'      00:01:00           0x23C34600, 600000000
'      01:00:00           0x61C46808, 1640261640
'      01:00:00.0000001   0x61C46809, 1640261641
'      01:00:00.0010000   0x61C48F18, 1640271640
'      01:00:01           0x625CFE88, 1650261640
'    1.00:00:00           0x2A69C0C9, 711573705
'    1.00:00:00.0000001   0x2A69C0C8, 711573704
'    1.00:00:00.0010000   0x2A69E7D9, 711583705
'    1.00:00:01           0x2B025649, 721573449
'  100.00:00:00           0x914F4E94, -1857073516
'  100.00:00:00.0010000   0x914F6984, -1857066620
'  100.00:00:01           0x91E7D814, -1847076844
'*/

