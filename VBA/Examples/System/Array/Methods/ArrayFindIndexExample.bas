Attribute VB_Name = "ArrayFindIndexExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 17 2024
'@LastModified February 17, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.findindex?view=netframework-4.8.1#examples

'@Dependencies
'   PredicateEndsWithSaurus.cls

Option Explicit

''
' The following code example demonstrates all three overloads of the FindIndex generic method.
' An array of strings is created, containing 8 dinosaur names, two of which (at positions 1 and 5)
' end with "saurus". The code example also defines a search predicate method named EndsWithSaurus,
' which accepts a string parameter and returns a Boolean value indicating whether the input string
' ends in "saurus".
'
' The FindIndex<T>(T[], Predicate<T>) method overload traverses the array from the beginning,
' passing each element in turn to the EndsWithSaurus method. The search stops when the
' EndsWithSaurus method returns true for the element at position 1.
'
' The FindIndex<T>(T[], Int32, Predicate<T>) method overload is used to search the array beginning
' at position 2 and continuing to the end of the array. It finds the element at position 5.
' Finally, the FindIndex<T>(T[], Int32, Int32, Predicate<T>) method overload is used to search the
' range of three elements beginning at position 2. It returns -1 because there are no dinosaur
' names in that range that end with "saurus".
''
Public Sub ArrayFindIndexExample()
    Dim dinosaurs As DotNetLib.Array
    Set dinosaurs = Arrays.CreateInitialize1D(VBString.GetType(), _
                        "Compsognathus", "Amargasaurus", "Oviraptor", _
                        "Velociraptor", "Deinonychus", "Dilophosaurus", _
                        "Gallimimus", "Triceratops")
    
    Debug.Print
    Dim dinosaur As Variant
    For Each dinosaur In dinosaurs
        Debug.Print dinosaur
    Next
    
    Dim pvtEndsWithSaurus As DotNetLib.Predicate
    Set pvtEndsWithSaurus = Predicate.Create(PredicateEndsWithSaurus)
    
    Debug.Print VBString.Format( _
                    VBString.Unescape("\nArray.FindIndex(dinosaurs, EndsWithSaurus): {0}"), _
                    Arrays.FindIndex(dinosaurs, pvtEndsWithSaurus))
    
    Debug.Print VBString.Format( _
                    VBString.Unescape("\nArray.FindIndex(dinosaurs, 2, EndsWithSaurus): {0}"), _
                    Arrays.FindIndex2(dinosaurs, 2, pvtEndsWithSaurus))
    
    Debug.Print VBString.Format( _
                    VBString.Unescape("\nArray.FindIndex(dinosaurs, 2, 3, EndsWithSaurus): {0}"), _
                     Arrays.FindIndex3(dinosaurs, 2, 3, pvtEndsWithSaurus))
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
'Array.FindIndex(dinosaurs, EndsWithSaurus): 1
'
'Array.FindIndex(dinosaurs, 2, EndsWithSaurus): 5
'
'Array.FindIndex(dinosaurs, 2, 3, EndsWithSaurus): -1
' */
