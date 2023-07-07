VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DayOfWeekHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@IgnoreModule MissingModuleAnnotation
'@Folder("DotNet.System")
'@PredeclaredId

Option Explicit

'@DefaultMember
'@Description("Converts the value of a DayOfWeek to string value.")
Public Property Get ToString(ByVal value As DotNetLib.DayOfWeek) As String
    Select Case value
        Case DayOfWeek_Sunday: ToString = "Sunday"
        Case DayOfWeek_Monday: ToString = "Monday"
        Case DayOfWeek_Tuesday: ToString = "Tuesday"
        Case DayOfWeek_Wednesday: ToString = "Wednesday"
        Case DayOfWeek_Thursday: ToString = "Thursday"
        Case DayOfWeek_Friday: ToString = "Friday"
        Case DayOfWeek_Saturday: ToString = "Saturday"
    End Select
End Property
