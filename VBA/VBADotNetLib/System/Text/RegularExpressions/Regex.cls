VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents an immutable regular expression."
'@Folder "VBADotNetLib.System.Text.RegularExpressions"
'@PredeclaredId
'@ModuleDescription("Represents an immutable regular expression.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v2.0 September 25, 2023
'@LastModified September 25, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference

'@Remarks
' Singleton class for Regex containing constructors and static members.

Option Explicit

Private Type TRegexSingleton
   Regex As DotNetLib.RegexSingleton
End Type

Private this As TRegexSingleton

Private Sub Class_Initialize()
   Set this.Regex = New DotNetLib.RegexSingleton
End Sub

Private Sub Class_Terminate()
   Set this.Regex = Nothing
End Sub

'@Description("Escapes a minimal set of characters (\, *, +, ?, |, {, [, (,), ^, $, ., #, and white space) by replacing them with their escape codes. This instructs the regular expression engine to interpret these characters literally rather than as metacharacters.")
Public Function Escape(ByVal str As String) As String
Attribute Escape.VB_Description = "Escapes a minimal set of characters (, *, +, ?, |, {, [, (,), ^, $, ., #, and white space) by replacing them with their escape codes. This instructs the regular expression engine to interpret these characters literally rather than as metacharacters."
   Escape = this.Regex.Escape(str)
End Function

'@Description("Converts any escaped characters in the input string.")
Public Function Unescape(ByVal str As String) As String
Attribute Unescape.VB_Description = "Converts any escaped characters in the input string."
   Unescape = this.Regex.Unescape(str)
End Function

