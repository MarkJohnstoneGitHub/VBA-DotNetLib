Attribute VB_Name = "CharIsSurrogatePairExample"
'@Folder("Examples.System.Char.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 28, 2023
'@LastModified December 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.char.issurrogatepair?view=netframework-4.8.1#system-char-issurrogatepair(system-string-system-int32)
' https://www.geeksforgeeks.org/c-sharp-char-issurrogatepairstring-int32-method/

Option Explicit

''
' Demonstrates the Char.IsSurrogatePair(String, Int32) Method
''
Public Sub CharIsSurrogatePairExample()
    On Error GoTo ErrorHandler
    ' calling Check() Method
    Call Check("1234", 3)
    Call Check("Tsunami", 3)
    Call Check("psyc0lo", 4)
    
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("a" & "\uD800" & "\uDC00" & "y")
    Call Check(s1, 1)
    
    Debug.Print
    Debug.Print "s is null"
    Call Check(VBA.vbNullString, 4)
    
Exit Sub
ErrorHandler:
    Debug.Print "Exception Thrown: ";
    Debug.Print Err.Description
End Sub

Private Sub Check(ByVal s As String, ByVal i As Long)
    ' checking condition
    ' using IsSurrogatePair() Method
    Dim val As Boolean
    val = Char.IsSurrogatePair(s, i)
    
    '/ checking
    If (val) Then
        Debug.Print VBAString.Format("String '{0}' contains " _
            & "Surrogate pairs at s[{1}] and s[{2}]", _
                                        s, i, i + 1)
        #If Not Mac Then
            Dim messageBoxText As String
            messageBoxText = VBAString.Format("String '{0}' contains " _
            & "Surrogate pairs at s[{1}] and s[{2}]", _
                                        s, i, i + 1)
            WinAPIUser32.MessageBoxW 0, StrPtr(messageBoxText), StrPtr("Char.IsSurrogatePair"), 0
        #End If
                                        
    Else
        Debug.Print VBAString.Format("String '{0}' does't contain any " _
                        & "Surrogate pairs at s[{1}] and s[{2}]", _
                                                    s, i, i + 1)
    End If
End Sub

'Output:
'String '1234' does't contain any Surrogate pairs at s[3] and s[4]
'String 'Tsunami' does't contain any Surrogate pairs at s[3] and s[4]
'String 'psyc0lo' does't contain any Surrogate pairs at s[4] and s[5]
'String 'aê??z' contains Surrogate pairs at s[1] and s[2]
'
's is null
'Exception Thrown: Value cannot be null.
'Parameter Name: s
