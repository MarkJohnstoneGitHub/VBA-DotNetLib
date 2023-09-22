VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents the standard input, output, and error streams for console applications."
'@Folder("VBADotNetLib.System")
'@PredeclaredId
'@ModuleDescription("Represents the standard input, output, and error streams for console applications.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 6, 2023
'@LastModified September 6, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime?view=netframework-4.8.1

'@Remarks
' Singleton class for Console

Option Explicit

Private Type TConsole
    ConsoleSingleton As DotNetLib.ConsoleSingleton
End Type

Private this As TConsole

Private Sub Class_Initialize()
    Set this.ConsoleSingleton = New DotNetLib.ConsoleSingleton
End Sub

Private Sub Class_Terminate()
    Set this.ConsoleSingleton = Nothing
End Sub

'--------------------------------------------------------------------
'Methods
'--------------------------------------------------------------------

'@Static
'@Description("Clears the console buffer and corresponding console window of display information.")
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the console buffer and corresponding console window of display information."
    this.ConsoleSingleton.Clear
End Sub

'@Static
'@Description("Writes the specified string value, followed by the current line terminator, to the standard output stream.")
Public Sub WriteLine(Optional ByRef value As String = VBA.vbNullString)
Attribute WriteLine.VB_Description = "Writes the specified string value, followed by the current line terminator, to the standard output stream."
    If value = VBA.vbNullString Then
        this.ConsoleSingleton.WriteLine
    Else
        this.ConsoleSingleton.WriteLine2 value
    End If
End Sub

'@Static
'@Description("Writes the specified string value, followed by the current line terminator, to the standard output stream.")
Public Sub WriteLine2(ByRef Format As String, ParamArray values() As Variant)
Attribute WriteLine2.VB_Description = "Writes the specified string value, followed by the current line terminator, to the standard output stream."
    Dim argsCount As Long
    If ((LBound(values) = 0) And (UBound(values) = -1)) Then
        argsCount = 0
    Else
        argsCount = UBound(values) + 1
    End If

    Select Case argsCount
        Case 0
            this.ConsoleSingleton.WriteLine2 Format
        Case 1
            this.ConsoleSingleton.WriteLine3 Format, values(0)
        Case 2
            this.ConsoleSingleton.WriteLine4 Format, values(0), values(1)
        Case 3
            this.ConsoleSingleton.WriteLine5 Format, values(0), values(1), values(2)
        Case 4
            this.ConsoleSingleton.WriteLine6 Format, values(0), values(1), values(2), values(3)
        Case Is > 4
            Dim args() As Variant
            args = values
            this.ConsoleSingleton.WriteLine7 Format, args
    End Select
End Sub
