Attribute VB_Name = "ArrayTrueForAllExample"
'@Folder("Examples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 14, 2024
'@LastModified February 14, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.trueforall?view=netframework-4.8.1#examples

Option Explicit

''
' The following example determines whether the last character of each element in a string array is
' a number. It creates two string arrays. The first array includes both strings that end with
' alphabetic characters and strings that end with numeric characters. The second array consists
' only of strings that end with numeric characters. The example also defines an EndWithANumber
' method whose signature matches the Predicate<T> delegate. The example passes each array to the
' TrueForAll method along with a delegate that represents the EndsWithANumber method.
''
Public Sub ArrayTrueForAllExample()
    Dim values1 As DotNetLib.Array
    Set values1 = Arrays.CreateInitialize1D(VBString.GetType(), "Y2K", "A2000", "DC2A6", "MMXIV", "0C3")
    Dim values2 As DotNetLib.Array
    Set values2 = Arrays.CreateInitialize1D(VBString.GetType(), "Y2", "A2000", "DC2A6", "MMXIV_0", "0C3")
    
    ' Assign the PredicateEndsWithANumber method to the Predicate delegate.
    Dim pvtEndsWithANumber As DotNetLib.Predicate
    Set pvtEndsWithANumber = Predicate.Create(PredicateEndsWithANumber)
    
    
    If (Arrays.TrueForAll(values1, pvtEndsWithANumber)) Then
        Debug.Print "All elements end with an integer."
    Else
        Debug.Print "Not all elements end with an integer."
    End If
    
    If (Arrays.TrueForAll(values2, pvtEndsWithANumber)) Then
        Debug.Print "All elements end with an integer."
    Else
        Debug.Print "Not all elements end with an integer."
    End If
End Sub

' The example displays the following output:
'       Not all elements end with an integer.
'       All elements end with an integer.
