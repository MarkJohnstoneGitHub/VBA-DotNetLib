Attribute VB_Name = "TestingDotNetLibDateTime"
Option Explicit

Private Sub TestDateTime()
    Dim xDate As DotNetLib.DateTime
    
    Dim myDate As Date
    myDate = Now()
    
    Dim myDateTime As DotNetLib.DateTime
    With New DotNetLib.DateTime
        Set myDateTime = .FromOADate(myDate)
    End With
    Debug.Print myDateTime.ToString
    Debug.Print myDateTime.ToString("D")
    Debug.Print "IsDaylightSavingTime: "; myDateTime.IsDaylightSavingTime
    
    Dim myUTCDateTime As DotNetLib.DateTime
    Set myUTCDateTime = myDateTime.ToUniversalTime
    Debug.Print myUTCDateTime.ToString
    Debug.Print myUTCDateTime.ToString("D")
    
    With New DotNetLib.DateTime
        Set xDate = .CreateFromDateTimeKind(2001, 4, 29, 10, 15, 0, 0, DateTimeKind_Utc)
    End With
    Debug.Print xDate.ToString
    Debug.Print xDate.ToString("D")
    
    Dim myLocaleDateTime As DotNetLib.DateTime
    
    Set myLocaleDateTime = xDate.ToLocalTime
    Debug.Print myLocaleDateTime.ToString("D")
    Debug.Print myLocaleDateTime.ToString()
    
    Dim t As DotNetLib.TimeSpan
    
    With New DotNetLib.TimeSpan
        Set t = .Create2(5, 5, 0, 0, 0)
    End With
    
    Set xDate = xDate.Addition(myLocaleDateTime, t)
    Debug.Print xDate.ToString
    
End Sub


Private Sub TestDateTime2()
    Dim myDateTime As DotNetLib.DateTime
    Set myDateTime = New DotNetLib.DateTime
    myDateTime.IsDaylightSavingTime
    Debug.Print
End Sub
