VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GregorianCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Singleton class for GregorianCalendar containing factory methods and static members."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Singleton class for GregorianCalendar containing factory methods and static members.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.gregoriancalendar?view=netframework-4.8.1

'@Remarks
' Singleton class for GregorianCalendar containing factory methods and static members.

Option Explicit

Private Type TGregorianCalendar
   GregorianCalendar As DotNetLib.GregorianCalendarSingleton
End Type

Private this As TGregorianCalendar

Private Sub Class_Initialize()
   Set this.GregorianCalendar = New DotNetLib.GregorianCalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.GregorianCalendar = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the GregorianCalendar class using the default or specified GregorianCalendarTypes value.")
Public Function Create(Optional ByVal pType As GregorianCalendarTypes = GregorianCalendarTypes.GregorianCalendarTypes_Localized) As DotNetLib.GregorianCalendar
Attribute Create.VB_Description = "Initializes a new instance of the GregorianCalendar class using the default or specified GregorianCalendarTypes value."
   Set Create = this.GregorianCalendar.Create(pType)
End Function

'@Static
'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.GregorianCalendar.CurrentEra()
End Property

'@Static
'@Description("Represents the current era. This field is constant.")
Public Property Get ADEra() As Long
Attribute ADEra.VB_Description = "Represents the current era. This field is constant."
   ADEra = this.GregorianCalendar.ADEra()
End Property

'@Static
'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.GregorianCalendar.Equals(objA, objB)
End Function

'@Static
'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByVal Calendar As DotNetLib.Calendar) As DotNetLib.Calendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.GregorianCalendar.ReadOnly(Calendar)
End Function
