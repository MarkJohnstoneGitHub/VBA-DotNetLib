Attribute VB_Name = "DateTimeMaxValueExample"
'@IgnoreModule VariableNotUsed
'@Folder("Examples.DateTime")
Option Explicit

'@Description("The following example instantiates a DateTime object by passing its constructor an Int64 value that represents a number of ticks. Before invoking the constructor, the example ensures that this value is greater than or equal to DateTime.MinValue.Ticks and less than or equal to DateTime.MaxValue.Ticks. If not, it throws an ArgumentOutOfRangeException.")
Public Sub DateTimeMaxValueField()
    ' Attempt to assign an out-of-range value to a DateTime constructor.
    Dim numberOfTicks As LongLong
    numberOfTicks = "9223372036854775807" 'Int64.MaxValue
    
    Dim validDate As DateTime
    ' Validate the value.
    If (numberOfTicks >= DateTime.MinValue.Ticks And numberOfTicks <= DateTime.MaxValue.Ticks) Then
        Set validDate = DateTime.CreateFromTicks(numberOfTicks)
    ElseIf (numberOfTicks < DateTime.MinValue.Ticks) Then
        Debug.Print numberOfTicks & " is less than " & DateTime.MinValue.Ticks & " ticks."
    Else
        Debug.Print numberOfTicks & " is greater than " & DateTime.MaxValue.Ticks & " ticks."
    End If
    
    ' The example displays the following output:
    '   9,223,372,036,854,775,807 is greater than 3,155,378,975,999,999,999 ticks.
End Sub


