Attribute VB_Name = "DateTimeMaxValueExample"
'@Folder "Examples.System.DateTime.Fields"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 9, 2023
'@LastModified January 7, 2024

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.maxvalue?view=netframework-4.8.1#examples

Option Explicit

'@Description("The following example instantiates a DateTime object by passing its constructor an Int64 value that represents a number of ticks. Before invoking the constructor, the example ensures that this value is greater than or equal to DateTime.MinValue.Ticks and less than or equal to DateTime.MaxValue.Ticks. If not, it throws an ArgumentOutOfRangeException.")
Public Sub DateTimeMaxValue()
Attribute DateTimeMaxValue.VB_Description = "The following example instantiates a DateTime object by passing its constructor an Int64 value that represents a number of ticks. Before invoking the constructor, the example ensures that this value is greater than or equal to DateTime.MinValue.Ticks and less than or equal to DateTime.MaxValue.Ticks. If not, it throws an ArgumentOutOfRangeException."
    ' Attempt to assign an out-of-range value to a DateTime constructor.
    Dim numberOfTicks As LongLong
    numberOfTicks = "9223372036854775807" 'Int64.MaxValue
    
    ' Validate the value.
    If (numberOfTicks >= DateTime.MinValue.Ticks And numberOfTicks <= DateTime.MaxValue.Ticks) Then
        Dim validDate As DotNetLib.DateTime
        Set validDate = DateTime.CreateFromTicks(numberOfTicks)
        Debug.Print validDate.ToString()
    ElseIf (numberOfTicks < DateTime.MinValue.Ticks) Then
        Debug.Print VBString.Format("{0:N0} is less than {1:N0} ticks.", _
                        numberOfTicks, _
                        DateTime.MinValue.Ticks)
    Else
        Debug.Print VBString.Format("{0:N0} is greater than {1:N0} ticks.", _
                                    numberOfTicks, _
                                    DateTime.MaxValue.Ticks)
    End If
End Sub

' The example displays the following output:
'   9,223,372,036,854,775,807 is greater than 3,155,378,975,999,999,999 ticks.
