VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DayOfWeekHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Rubberduck annotations
'@Folder "VBADotNetLib.System"
'@PredeclaredId

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 09, 2023
'@LastModified August 4, 2023

'@ReferenceAddin DotNetLib.tlb

Option Explicit

'@DefaultMember
'@Description("Converts the value of a DayOfWeek to string value.")
Public Property Get ToString(ByVal value As DotNetLib.DayOfWeek) As String
Attribute ToString.VB_Description = "Converts the value of a DayOfWeek to string value."
Attribute ToString.VB_UserMemId = 0
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
