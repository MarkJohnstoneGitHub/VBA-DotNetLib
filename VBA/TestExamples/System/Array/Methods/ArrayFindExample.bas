Attribute VB_Name = "ArrayFindExample"
'@Folder("TestExamples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 8, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.find?view=netframework-4.8.1

Option Explicit

Public Sub ArrayFindExample()
    ' Create and initialize a new array.
    Dim words As DotNetLib.Array
    Set words = Arrays.CreateInitialize1D(VBString.GetType(), _
                        "The", "Magpie", "Bobbie", "Bob the great", "Sacred Bob", "FOX", "jumps", _
                        "over", "the", "lazy", "dog")

    ' Find the first array item containing the word "bob" ignore case
    
    ' Assign the PredicateContainsBob method to the Predicate delegate.
    Dim pvtPredicate As DotNetLib.Predicate
    Set pvtPredicate = Predicate.Create(PredicateContainsBob)
    
    Dim pvtFirst As String
    pvtFirst = Arrays.Find(words, pvtPredicate)
    Debug.Print VBString.Format("Found: {0}", pvtFirst)
    Debug.Print
    
    Dim pvtIndex As Long
    pvtIndex = Arrays.FindIndex(words, pvtPredicate)
    Debug.Print VBString.Format("Found at index: {0}", pvtIndex)
    Debug.Print
    
    Dim pvtResult As DotNetLib.Array
    Set pvtResult = Arrays.FindAll(words, pvtPredicate)
    Dim varItem As Variant
    For Each varItem In pvtResult
        Debug.Print varItem
    Next
End Sub


Public Sub ArrayFindExampleV2()
    ' Create and initialize a new array.
    Dim words As DotNetLib.Array
    Set words = Arrays.CreateInitialize1D(VBString.GetType(), _
                        "The", "Magpie", "bobbie", "Bob the great", "Sacred Bob", "FOX", "jumps", _
                        "over", "The awesome Rubberduck", "lazy", "dog")

    ' Find the first array item containing the word "the" ignore case
    ' Assign the PredicateContainsString method to the Predicate delegate.
    Dim pvtPredicateContainsString As DotNetLib.Predicate
    Set pvtPredicateContainsString = PredicateContainsString.Create("the")
    
    Dim pvtFirst As String
    pvtFirst = Arrays.Find(words, pvtPredicateContainsString)
    Debug.Print VBString.Format("Found: {0}", pvtFirst)
    Debug.Print
    
    Dim pvtIndex As Long
    pvtIndex = Arrays.FindIndex(words, pvtPredicateContainsString)
    Debug.Print VBString.Format("Found at index: {0}", pvtIndex)
    Debug.Print

    Dim pvtResult As DotNetLib.Array
    Set pvtResult = Arrays.FindAll(words, pvtPredicateContainsString)
    Dim varItem As Variant
    For Each varItem In pvtResult
        Debug.Print varItem
    Next
End Sub
