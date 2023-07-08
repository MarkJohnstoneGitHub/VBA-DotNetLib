Attribute VB_Name = "DateTimeAdd"
'@Folder("Examples.DateTime.Methods")

' https://learn.microsoft.com/en-us/dotnet/api/system.datetime.add?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example demonstrates the Add method. It calculates the day of the week that is 36 days (864 hours) from this moment.")
Public Sub DateTimeAdd()
    Dim today As DateTime
    Set today = DateTime.Now
    
    Dim duration As DotNetLib.TimeSpan
    With New DotNetLib.TimeSpan
        Set duration = .Create2(36, 0, 0, 0)
    End With
    
    Dim answer As DateTime
    Set answer = today.Add(duration)
    Debug.Print answer.ToString()

End Sub

