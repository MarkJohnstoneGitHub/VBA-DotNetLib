Attribute VB_Name = "HashtableContainsExample"
'@Folder("Examples.System.Collections.Hashtable.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 22, 2023
'@LastModified October 22, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable.contains?view=netframework-4.8.1#examples

'@Remarks
' Identifier type characters
' https://learn.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/data-types/type-characters#identifier-type-characters
' &   Long
Option Explicit

''
' The following example shows how to determine whether the Hashtable contains
' a specific element.
''
Public Sub HashtableContains()
    ' Creates and initializes a new Hashtable.
    Dim myHT As DotNetLib.Hashtable
    Set myHT = Hashtable.Create()
    'Note &  Type Identifier Character for Long
    Call myHT.Add(0&, "zero")
    Call myHT.Add(1&, "one")
    Call myHT.Add(2&, "two")
    Call myHT.Add(3&, "three")
    Call myHT.Add(4&, "four")
    
    ' Displays the values of the Hashtable.
    Debug.Print "The Hashtable contains the following values:"
    Call PrintIndexAndKeysAndValues(myHT)
    
    ' Searches for a specific key.
    Dim myKey As Long
    myKey = 2
    Debug.Print VBAString.Format("The key ""{0}"" is {1}.", myKey, IIf(myHT.ContainsKey(myKey), "in the Hashtable", "NOT in the Hashtable"))
    myKey = 6
    Debug.Print VBAString.Format("The key ""{0}"" is {1}.", myKey, IIf(myHT.ContainsKey(myKey), "in the Hashtable", "NOT in the Hashtable"))

    ' Searches for a specific value.
    Dim myValue As Variant
    myValue = "three"
    Debug.Print VBAString.Format("The value ""{0}"" is {1}.", myValue, IIf(myHT.ContainsValue(myValue), "in the Hashtable", "NOT in the Hashtable"))
    myValue = "nine"
    Debug.Print VBAString.Format("The value ""{0}"" is {1}.", myValue, IIf(myHT.ContainsValue(myValue), "in the Hashtable", "NOT in the Hashtable"))
End Sub

Private Sub PrintIndexAndKeysAndValues(ByVal myHT As DotNetLib.Hashtable)
    Dim formatString As String
    formatString = Regex.Unescape("\t[{0}]:\t{1}\t{2}")
    Debug.Print Regex.Unescape("\t-INDEX-\t-KEY-\t-VALUE-")
    Dim i As Long
    i = 0
    Dim varDictEntry As Variant
    For Each varDictEntry In myHT
        Dim dictEntry As mscorlib.DictionaryEntry
        dictEntry = varDictEntry
        Debug.Print VBAString.Format(formatString, i, DictionaryEntry.Key(dictEntry), DictionaryEntry.value(dictEntry))
        i = i + 1
    Next
End Sub

'/*
'This code produces the following output.
'
'The Hashtable contains the following values:
'        -INDEX- -KEY-   -VALUE-
'        [0]:    4       four
'        [1]:    3       three
'        [2]:    2       two
'        [3]:    1       one
'        [4]:    0       zero
'
'The key "2" is in the Hashtable.
'The key "6" is NOT in the Hashtable.
'The value "three" is in the Hashtable.
'The value "nine" is NOT in the Hashtable.
'
'*/
