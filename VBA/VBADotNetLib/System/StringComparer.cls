VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a string comparison operation that uses specific case and culture-based or ordinal comparison rules."
'@Folder "VBADotNetLib.System"
'@ModuleDescription("Represents a string comparison operation that uses specific case and culture-based or ordinal comparison rules.")
'@PredeclaredId

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 17, 2023
'@LastModified October 17, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.stringcomparer?view=netframework-4.8.1

'@Remarks
' Singleton class for StringComparer containing constructors and static members.

Option Explicit

Private Type TStringComparerSingleton
   StringComparerSingleton As DotNetLib.StringComparerSingleton
End Type

Private this As TStringComparerSingleton

Private Sub Class_Initialize()
   Set this.StringComparerSingleton = New DotNetLib.StringComparerSingleton
End Sub

Private Sub Class_Terminate()
   Set this.StringComparerSingleton = Nothing
End Sub

'@Description("Gets a StringComparer object that performs a case-sensitive string comparison using the word comparison rules of the current culture.")
Public Property Get CurrentCulture() As DotNetLib.StringComparer
Attribute CurrentCulture.VB_Description = "Gets a StringComparer object that performs a case-sensitive string comparison using the word comparison rules of the current culture."
   Set CurrentCulture = this.StringComparerSingleton.CurrentCulture()
End Property

'@Description("Gets a StringComparer object that performs case-insensitive string comparisons using the word comparison rules of the current culture.")
Public Property Get CurrentCultureIgnoreCase() As DotNetLib.StringComparer
Attribute CurrentCultureIgnoreCase.VB_Description = "Gets a StringComparer object that performs case-insensitive string comparisons using the word comparison rules of the current culture."
   Set CurrentCultureIgnoreCase = this.StringComparerSingleton.CurrentCultureIgnoreCase()
End Property

'@Description("Gets a StringComparer object that performs a case-sensitive string comparison using the word comparison rules of the invariant culture.")
Public Property Get InvariantCulture() As DotNetLib.StringComparer
Attribute InvariantCulture.VB_Description = "Gets a StringComparer object that performs a case-sensitive string comparison using the word comparison rules of the invariant culture."
   Set InvariantCulture = this.StringComparerSingleton.InvariantCulture()
End Property

'@Description("Gets a StringComparer object that performs a case-insensitive string comparison using the word comparison rules of the invariant culture.")
Public Property Get InvariantCultureIgnoreCase() As DotNetLib.StringComparer
Attribute InvariantCultureIgnoreCase.VB_Description = "Gets a StringComparer object that performs a case-insensitive string comparison using the word comparison rules of the invariant culture."
   Set InvariantCultureIgnoreCase = this.StringComparerSingleton.InvariantCultureIgnoreCase()
End Property

'@Description("Gets a StringComparer object that performs a case-sensitive ordinal string comparison.")
Public Property Get Ordinal() As DotNetLib.StringComparer
Attribute Ordinal.VB_Description = "Gets a StringComparer object that performs a case-sensitive ordinal string comparison."
   Set Ordinal = this.StringComparerSingleton.Ordinal()
End Property

'@Description("Gets a StringComparer object that performs a case-insensitive ordinal string comparison.")
Public Property Get OrdinalIgnoreCase() As DotNetLib.StringComparer
Attribute OrdinalIgnoreCase.VB_Description = "Gets a StringComparer object that performs a case-insensitive ordinal string comparison."
   Set OrdinalIgnoreCase = this.StringComparerSingleton.OrdinalIgnoreCase()
End Property

'@Description("Creates a StringComparer object that compares strings according to the rules of a specified culture.")
Public Function Create(ByVal culture As DotNetLib.CultureInfo, ByVal ignoreCase As Boolean) As DotNetLib.StringComparer
Attribute Create.VB_Description = "Creates a StringComparer object that compares strings according to the rules of a specified culture."
   Set Create = this.StringComparerSingleton.Create(culture, ignoreCase)
End Function
