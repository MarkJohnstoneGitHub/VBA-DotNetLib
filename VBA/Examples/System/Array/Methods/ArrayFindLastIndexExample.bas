Attribute VB_Name = "ArrayFindLastIndexExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 14, 2024
'@LastModified February 14, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.findlastindex?view=netframework-4.8.1#examples

'@Dependencies PredicateEndsWithSaurus.cls

Option Explicit

''
' The following code example demonstrates all three overloads of the FindLastIndex generic method.
' An array of strings is created, containing 8 dinosaur names, two of which (at positions 1 and 5)
' end with "saurus". The code example also defines a search predicate method named EndsWithSaurus,
' which accepts a string parameter and returns a Boolean value indicating whether the input string
' ends in "saurus".
'
' The FindLastIndex<T>(T[], Predicate<T>) method overload traverses the array backward from the end,
' passing each element in turn to the EndsWithSaurus method. The search stops when the
' EndsWithSaurus method returns true for the element at position 5.
'
' The FindLastIndex<T>(T[], Int32, Predicate<T>) method overload is used to search the array
' beginning at position 4 and continuing backward to the beginning of the array. It finds the
' element at position 1. Finally, the FindLastIndex<T>(T[], Int32, Int32, Predicate<T>) method
' overload is used to search the range of three elements beginning at position 4 and working
' backward (that is, elements 4, 3, and 2). It returns -1 because there are no dinosaur names in
' that range that end with "saurus".
''
Public Sub ArrayFindLastIndexExample()
    Dim dinosaurs As DotNetLib.Array
    Set dinosaurs = Arrays.CreateInitialize1D(VBString.GetType(), _
                        "Compsognathus", "Amargasaurus", "Oviraptor", "Velociraptor", _
                        "Deinonychus", "Dilophosaurus", "Gallimimus", _
                        "Triceratops")
                        
    Debug.Print
    Dim dinosaur As Variant
    For Each dinosaur In dinosaurs
        Debug.Print dinosaur
    Next
    Dim pvtEndsWithSaurus As DotNetLib.Predicate
    Set pvtEndsWithSaurus = Predicate.Create(PredicateEndsWithSaurus)
    Debug.Print VBString.Format( _
            VBString.Unescape("\nArray.FindLastIndex(dinosaurs, EndsWithSaurus): {0}"), _
            Arrays.FindLastIndex(dinosaurs, pvtEndsWithSaurus))
            
    Debug.Print VBString.Format( _
            VBString.Unescape("\nArray.FindLastIndex(dinosaurs, 4, EndsWithSaurus): {0}"), _
            Arrays.FindLastIndex2(dinosaurs, 4, pvtEndsWithSaurus))
            
    Debug.Print VBString.Format( _
            VBString.Unescape("\nArray.FindLastIndex(dinosaurs, 4, 3, EndsWithSaurus): {0}"), _
            Arrays.FindLastIndex3(dinosaurs, 4, 3, pvtEndsWithSaurus))
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
'Array.FindLastIndex(dinosaurs, EndsWithSaurus): 5
'
'Array.FindLastIndex(dinosaurs, 4, EndsWithSaurus): 1
'
'Array.FindLastIndex(dinosaurs, 4, 3, EndsWithSaurus): -1
' */

