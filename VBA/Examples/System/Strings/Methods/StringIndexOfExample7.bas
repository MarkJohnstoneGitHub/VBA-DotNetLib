Attribute VB_Name = "StringIndexOfExample7"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 31, 2023
'@LastModified December 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.indexof?view=netframework-4.8.1#system-string-indexof(system-string-system-int32)

Option Explicit

''
' The following example searches for all occurrences of a specified string
' within a target string.
''
Public Sub StringIndexOfExample7()
    Dim strSource As DotNetLib.String
    Set strSource = Strings.Create("This is the string which we will perform the search on")
    
    Debug.Print VBAString.Format("The search string is:{0}""{1}""{0}", Environment.NewLine, strSource)

    Dim strTarget As String
    strTarget = VBA.vbNullString
    Dim found As Long
    Dim totFinds As Long
    Do
        strTarget = Strings.Create(InputBox(VBAString.Format("The search string is:{0}""{1}""{0}", Environment.NewLine, strSource) + "Please enter a search value to look for in the above string (hit Enter to exit) ==> "))
        If (strTarget <> "") Then
            Dim i As Long
            For i = 0 To strSource.length - 1
                found = strSource.IndexOf9(strTarget, i)
                If (found >= 0) Then
                    totFinds = totFinds + 1
                    i = found
                Else
                    Exit For
                End If
            Next
        Else
            Exit Do
        End If
        Debug.Print VBAString.Format("{0}The search parameter '{1}' was found {2} times.{0}", _
                    Environment.NewLine, strTarget, totFinds)
        totFinds = 0
    Loop While (True)
End Sub
