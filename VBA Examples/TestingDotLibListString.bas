<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
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
=======
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
>>>>>>> 7342d40c8d06560e9faa22506c24db1b9cd2eb7f
=======
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
>>>>>>> 7342d40c8d06560e9faa22506c24db1b9cd2eb7f
=======
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
>>>>>>> 7342d40c8d06560e9faa22506c24db1b9cd2eb7f
