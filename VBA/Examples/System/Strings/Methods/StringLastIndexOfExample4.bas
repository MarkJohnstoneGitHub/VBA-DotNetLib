Attribute VB_Name = "StringLastIndexOfExample4"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 2, 2024
'@LastModified January 2, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-string-system-int32-system-int32)

Option Explicit

''
' In the following example, the LastIndexOf method is used to find the position
' of a soft hyphen (U+00AD) followed by an "m" or "n" in two strings. Only one
' of the strings contains a soft hyphen. In the case of the string that includes
' the soft hyphen followed by an "m", LastIndexOf returns the index of the "m"
' when searching for the soft hyphen followed by "m".
''
Public Sub StringLastIndexOfExample4()
    Dim position As Long
    position = 0
    Dim s1 As DotNetLib.String
    Set s1 = Strings.CreateUnescape("ani\u00ADmal")
    Dim s2 As DotNetLib.String
    Set s2 = Strings.Create("animal")
    
    ' Find the index of the soft hyphen followed by "n".
    position = s1.LastIndexOf7("m")
    Debug.Print VBString.Format("'m' at position {0}", position)

    If (position >= 0) Then
        Debug.Print VBString.Format(s1.LastIndexOf3(Strings.CreateUnescape("\u00ADn"), position, position + 1))
    End If
    
    position = s2.LastIndexOf7("m")
    Debug.Print VBString.Format("'m' at position {0}", position)


    'if (position >= 0)
    '    Console.WriteLine(s2.LastIndexOf("\u00ADn", position, position + 1));

    If (position >= 0) Then
        Debug.Print VBString.Format(s2.LastIndexOf3(Strings.CreateUnescape("\u00ADn"), position, position + 1))
    End If

    ' Find the index of the soft hyphen followed by "m".
    position = s1.LastIndexOf7("m")
    Debug.Print VBString.Format("'m' at position {0}", position)

    If (position >= 0) Then
        Debug.Print VBString.Format(s1.LastIndexOf3(Strings.CreateUnescape("\u00ADm"), position, position + 1))
    End If
    
    position = s2.LastIndexOf7("m")
    Debug.Print VBString.Format("'m' at position {0}", position)
    
    If (position >= 0) Then
        Debug.Print VBString.Format(s2.LastIndexOf3(Strings.CreateUnescape("\u00ADm"), position, position + 1));
    End If
End Sub

' The example displays the following output:
'
' 'm' at position 4
' 1
' 'm' at position 3
' 1
' 'm' at position 4
' 4
' 'm' at position 3
' 3
