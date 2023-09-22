VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UmAlQuraCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the Saudi Hijri (Um Al Qura) calendar."
'@Folder "VBADotNetLib.System.Globalization"
'@PredeclaredId
'@ModuleDescription("Represents the Saudi Hijri (Um Al Qura) calendar.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 19, 2023
'@LastModified September 22, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.umalquracalendar?view=netframework-4.8.1

'@Remarks
' Singleton class for UmAlQuraCalendar containing factory methods and static members.

Option Explicit

Private Type TUmAlQuraCalendar
   UmAlQuraCalendar As DotNetLib.UmAlQuraCalendarSingleton
End Type

Private this As TUmAlQuraCalendar

Private Sub Class_Initialize()
   Set this.UmAlQuraCalendar = New DotNetLib.UmAlQuraCalendarSingleton
End Sub

Private Sub Class_Terminate()
   Set this.UmAlQuraCalendar = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the UmAlQuraCalendar class.")
Public Function Create() As DotNetLib.UmAlQuraCalendar
Attribute Create.VB_Description = "Initializes a new instance of the UmAlQuraCalendar class."
   Set Create = this.UmAlQuraCalendar.Create()
End Function

'@Static
'@Description("Represents the current era of the current calendar. The value of this field is 0.")
Public Property Get CurrentEra() As Long
Attribute CurrentEra.VB_Description = "Represents the current era of the current calendar. The value of this field is 0."
   CurrentEra = this.UmAlQuraCalendar.CurrentEra()
End Property

'@Static
'@Description("Represents the current era. This field is constant.")
Public Property Get UmAlQuraEra() As Long
Attribute UmAlQuraEra.VB_Description = "Represents the current era. This field is constant."
   UmAlQuraEra = this.UmAlQuraCalendar.UmAlQuraEra()
End Property

'@Static
'@Description("Determines whether the specified object instances are considered equal.")
Public Function Equals(ByVal objA As Variant, ByVal objB As Variant) As Boolean
Attribute Equals.VB_Description = "Determines whether the specified object instances are considered equal."
   Equals = this.UmAlQuraCalendar.Equals(objA, objB)
End Function

'@Static
'@Description("Returns a read-only version of the specified Calendar object.")
Public Function ReadOnly(ByVal Calendar As DotNetLib.Calendar) As DotNetLib.Calendar
Attribute ReadOnly.VB_Description = "Returns a read-only version of the specified Calendar object."
   Set ReadOnly = this.UmAlQuraCalendar.ReadOnly(Calendar)
End Function
