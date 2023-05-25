Attribute VB_Name = "TestingDotLibListString"
'@Folder("Database2")
Option Explicit

Private Sub TestDotNetLibListString()
    Dim dinosaurs As DotNetLib.ListString

    With New DotNetLib.ListString
        Set dinosaurs = .Create
    End With
    
    dinosaurs.Add "Tyrannosaurus"
    dinosaurs.Add ("Amargasaurus")
    dinosaurs.Add ("Deinonychus")
    dinosaurs.Add ("Compsognathus")
    
    dinosaurs.Sort
    
    Debug.Print dinosaurs.Contains("Amargasaurus")
    dinosaurs.CopyTo
    
    
    Dim arrayDinosaurs() As String
    arrayDinosaurs = dinosaurs.ToArray

End Sub
