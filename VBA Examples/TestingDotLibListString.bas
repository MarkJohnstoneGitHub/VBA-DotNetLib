Attribute VB_Name = "TestingDotLibListString"
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
    
    Dim arrayDinosaurs() As String
    arrayDinosaurs = dinosaurs.ToArray

End Sub
