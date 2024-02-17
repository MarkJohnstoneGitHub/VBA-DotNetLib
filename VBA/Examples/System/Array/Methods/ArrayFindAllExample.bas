Attribute VB_Name = "ArrayFindAllExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 17 2024
'@LastModified February 17, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.findall?view=netframework-4.8.1#examples

'@Dependencies
'   PredicateEndsWithRaptor.cls
'   PredicateEndsWithSaurus.cls
'   PredicateEndsWithTops.cls
'
Option Explicit

Private dinosaurs As DotNetLib.Array

''
' The following code example demonstrates the Find, FindLast, and FindAll generic methods. An array
' of strings is created, containing 8 dinosaur names, two of which (at positions 1 and 5) end with
' "saurus". The code example also defines a search predicate method named EndsWithSaurus, which
' accepts a string parameter and returns a Boolean value indicating whether the input string ends
' in "saurus".
'
' The Find generic method traverses the array from the beginning, passing each element in turn to
' the EndsWithSaurus method. The search stops when the EndsWithSaurus method returns true for the
' element "Amargasaurus".
'
' The FindLast generic method is used to search the array backward from the end. It finds the
' element "Dilophosaurus" at position 5. The FindAll generic method is used to return an array
' containing all the elements that end in "saurus". The elements are displayed.
'
' The code example also demonstrates the Exists and TrueForAll generic methods.
''
Public Sub ArrayFindAllExample()
    Set dinosaurs = Arrays.CreateInitialize1D(VBString.GetType(), _
                        "Compsognathus", "Amargasaurus", "Oviraptor", _
                        "Velociraptor", "Deinonychus", "Dilophosaurus", _
                        "Gallimimus", "Triceratops")
                        
    Call DiscoverAll
    Call DiscoveryByEnding("saurus")
End Sub

Private Sub DiscoverAll()
    Debug.Print
    Dim dinosaur As Variant
    For Each dinosaur In dinosaurs
        Debug.Print dinosaur
    Next
End Sub

Private Sub DiscoveryByEnding(ByVal ending As String)
    Dim dinoType As DotNetLib.Predicate
    Select Case LCase$(ending)
        Case "raptor"
            Set dinoType = Predicate.Create(PredicateEndsWithRaptor)
        Case "tops"
            Set dinoType = Predicate.Create(PredicateEndsWithTops)
        Case "saurus"
            Set dinoType = Predicate.Create(PredicateEndsWithSaurus)
    End Select
    
    Debug.Print VBString.Format( _
            VBString.Unescape("\nArray.Exists(dinosaurs, ""{0}""): {1}"), _
            ending, _
            Arrays.Exists(dinosaurs, dinoType))

    Debug.Print VBString.Format( _
                    VBString.Unescape("\nArray.TrueForAll(dinosaurs, ""{0}""): {1}"), _
                    ending, _
                    Arrays.TrueForAll(dinosaurs, dinoType))
    
    Debug.Print VBString.Format( _
                    VBString.Unescape("\nArray.Find(dinosaurs, ""{0}""): {1}"), _
                    ending, _
                    Arrays.Find(dinosaurs, dinoType))
            
    Debug.Print VBString.Format( _
                    VBString.Unescape("\nArray.FindLast(dinosaurs, ""{0}""): {1}"), _
                     ending, _
                     Arrays.FindLast(dinosaurs, dinoType))
                     
    Debug.Print VBString.Format( _
             VBString.Unescape("\nArray.FindAll(dinosaurs, ""{0}""):"), ending)

    Dim subArray As DotNetLib.Array
    Set subArray = Arrays.FindAll(dinosaurs, dinoType)

    Dim dinosaur As Variant
    For Each dinosaur In subArray
        Debug.Print dinosaur
    Next
End Sub

'/* This code example produces the following output:
'
'Compsognathus
'Amargasaurus
'Oviraptor
'Velociraptor
'Deinonychus
'Dilophosaurus
'Gallimimus
'Triceratops
'
'Array.Exists(dinosaurs, "saurus"): True
'
'Array.TrueForAll(dinosaurs, "saurus"): False
'
'Array.Find(dinosaurs, "saurus"): Amargasaurus
'
'Array.FindLast(dinosaurs, "saurus"): Dilophosaurus
'
'Array.FindAll(dinosaurs, "saurus"):
'Amargasaurus
'Dilophosaurus
'*/
