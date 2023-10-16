VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CaseInsensitiveComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Compares two objects for equivalence, ignoring the case of strings."
'@Folder "VBADotNetLib.System.Collections"
'@ModuleDescription("Compares two objects for equivalence, ignoring the case of strings.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 17, 2023
'@LastModified October 17, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.caseinsensitivecomparer?view=netframework-4.8.1

'@Remarks
'Singleton class for CaseInsensitiveComparer containing factory methods and static members.


Option Explicit

Private Type TCaseInsensitiveComparerSingleton
   CaseInsensitiveComparerSingleton As DotNetLib.CaseInsensitiveComparerSingleton
End Type

Private this As TCaseInsensitiveComparerSingleton

Private Sub Class_Initialize()
   Set this.CaseInsensitiveComparerSingleton = New DotNetLib.CaseInsensitiveComparerSingleton
End Sub

Private Sub Class_Terminate()
   Set this.CaseInsensitiveComparerSingleton = Nothing
End Sub

'@Static
'@Description("Initializes a new instance of the CaseInsensitiveComparer class using the CurrentCulture of the current thread.")
Public Function Create() As DotNetLib.CaseInsensitiveComparer
Attribute Create.VB_Description = "Initializes a new instance of the CaseInsensitiveComparer class using the CurrentCulture of the current thread."
   Set Create = this.CaseInsensitiveComparerSingleton.Create()
End Function

'@Static
'@Description("Initializes a new instance of the CaseInsensitiveComparer class using the specified CultureInfo.")
Public Function Create2(ByRef culture As DotNetLib.CultureInfo) As DotNetLib.CaseInsensitiveComparer
Attribute Create2.VB_Description = "Initializes a new instance of the CaseInsensitiveComparer class using the specified CultureInfo."
   Set Create2 = this.CaseInsensitiveComparerSingleton.Create2(culture)
End Function

'@Static
'@Description("Gets an instance of CaseInsensitiveComparer that is associated with the CurrentCulture of the current thread and that is always available.")
Public Property Get Default() As DotNetLib.CaseInsensitiveComparer
Attribute Default.VB_Description = "Gets an instance of CaseInsensitiveComparer that is associated with the CurrentCulture of the current thread and that is always available."
   Set Default = this.CaseInsensitiveComparerSingleton.Default()
End Property

'@Static
'@Description("Gets an instance of CaseInsensitiveComparer that is associated with InvariantCulture and that is always available.")
Public Property Get DefaultInvariant() As DotNetLib.CaseInsensitiveComparer
Attribute DefaultInvariant.VB_Description = "Gets an instance of CaseInsensitiveComparer that is associated with InvariantCulture and that is always available."
   Set DefaultInvariant = this.CaseInsensitiveComparerSingleton.DefaultInvariant()
End Property
