Attribute VB_Name = "ChangeDirectory"
'@Folder("Testing.Junk")

Option Explicit

''@Description("Change the current directory.")
''@Exceptions
'   VBAErrorFileNotFound 76
'Private Sub ChangeCurrentDirectory(ByVal path As String)
'    If Not DirectoryExist(path) Then
'        Err.Raise VBAErrorFileNotFound, "RubberduckUtility.ExportAllComponents", "Output directory not found: " & path
'    End If
'
'    Dim folders As Variant
'    folders = Split(path, DirectorySeparatorChar)
'    Dim drive As String
'    drive = folders(0)
'    ChDrive drive
'    ChDir path
'End Sub
