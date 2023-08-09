Attribute VB_Name = "TimeSpanNegateExample"
'@Folder("VBADotNetLib.Examples.TimeSpan.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 17, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.timespan.negate?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example applies the Negate method to several TimeSpan objects.")
Public Sub TimeSpanNegate()
Attribute TimeSpanNegate.VB_Description = "The following example applies the Negate method to several TimeSpan objects."
   Debug.Print "TimeSpan        ", "Duration( )        ", "Negate( )"
   Debug.Print "--------        ", "---------          ", "---------"
   
   ' Create TimeSpan objects and apply the Unary Negation
   ' and Unary Plus operators to them.
   ShowDurationNegate TimeSpan.CreateFromTicks(1)
   ShowDurationNegate TimeSpan.CreateFromTicks(-1234567)
   ShowDurationNegate TimeSpan.UnaryPlus(TimeSpan.Create3(0, 0, 10, -20, -30))
   ShowDurationNegate TimeSpan.UnaryPlus(TimeSpan.Create3(0, -10, 20, -30, 40))
   ShowDurationNegate TimeSpan.UnaryNegation(TimeSpan.Create3(1, 10, 20, 40, 160))
   ShowDurationNegate TimeSpan.UnaryNegation(TimeSpan.Create3(-10, -20, -30, -40, -50))
End Sub

Private Sub ShowDurationNegate(ByVal interval As ITimeSpan)
   ' Display the TimeSpan value and the results of the
   ' Duration and Negate methods.
   Debug.Print interval.ToString(), interval.Duration.ToString(), interval.Negate.ToString()
End Sub

'/*
'This example of TimeSpan.Duration( ), TimeSpan.Negate( ),
'and the TimeSpan Unary Negation and Unary Plus operators
'generates the following output.
'
'              TimeSpan           Duration( )             Negate( )
'              --------           -----------             ---------
'      00:00:00.0000001      00:00:00.0000001     -00:00:00.0000001
'     -00:00:00.1234567      00:00:00.1234567      00:00:00.1234567
'      00:09:39.9700000      00:09:39.9700000     -00:09:39.9700000
'     -09:40:29.9600000      09:40:29.9600000      09:40:29.9600000
'   -1.10:20:40.1600000    1.10:20:40.1600000    1.10:20:40.1600000
'   10.20:30:40.0500000   10.20:30:40.0500000  -10.20:30:40.0500000
'*/
