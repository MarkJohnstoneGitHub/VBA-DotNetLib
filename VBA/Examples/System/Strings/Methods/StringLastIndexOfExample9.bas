Attribute VB_Name = "StringLastIndexOfExample9"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 3, 2024
'@LastModified January 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-char)

Option Explicit

''
' The following example defines an ExtractFilename method that uses the
' LastIndexOf(Char) method to find the last directory separator character in a
' string and to extract the string's file name. If the file exists, the method
' returns the file name without its path.
''
Public Sub StringLastIndexOfExample9()
    Dim filename As DotNetLib.String

    Set filename = ExtractFilename(Strings.Create("C:\temp\"))
    Debug.Print VBAString.Format("{0}", IIf(Strings.IsNullOrEmpty(filename), "<none>", filename))

    Set filename = ExtractFilename(Strings.Create("C:\temp\delegate.txt"))
    Debug.Print VBAString.Format("{0}", IIf(Strings.IsNullOrEmpty(filename), "<none>", filename))

    Set filename = ExtractFilename(Strings.Create("delegate.txt"))
    Debug.Print VBAString.Format("{0}", IIf(Strings.IsNullOrEmpty(filename), "<none>", filename))

    Set filename = ExtractFilename(Strings.Create("C:\temp\notafile.txt"))
    Debug.Print VBAString.Format("{0}", IIf(Strings.IsNullOrEmpty(filename), "<none>", filename))
End Sub

Private Function ExtractFilename(ByVal filepath As DotNetLib.String) As DotNetLib.String
    ' If path ends with a "\", it's a path only so return String.Empty.
    If (filepath.Trim().EndsWith3("\")) Then
        Set ExtractFilename = Strings.EmptyString()
        Exit Function
    End If
    
    ' Determine where last backslash is.
    Dim position As Long
    position = filepath.LastIndexOf7("\")
    ' If there is no backslash, assume that this is a filename.
    If (position = -1) Then
        ' Determine whether file exists in the current directory.
        If (File.Exists(Environment.CurrentDirectory + Path.DirectorySeparatorChar + filepath)) Then
            Set ExtractFilename = filepath
            Exit Function
        Else
            Set ExtractFilename = Strings.EmptyString()
            Exit Function
        End If
    Else
        ' Determine whether file exists using filepath.
        If (File.Exists(filepath)) Then
            ' Return filename without file path.
            Set ExtractFilename = filepath.substring(position + 1)
            Exit Function
        Else
            Set ExtractFilename = Strings.EmptyString()
            Exit Function
        End If
    End If
End Function

