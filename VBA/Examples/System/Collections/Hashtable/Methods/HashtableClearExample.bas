Attribute VB_Name = "HashtableClearExample"
'@Folder("Examples.System.Collections.Hashtable.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 22, 2023
'@LastModified October 23, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable.clear?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to clear the values of the Hashtable.
''
Public Sub HashtableClear()
    ' Creates and initializes a new Hashtable.
    Dim myHT As DotNetLib.Hashtable
    Set myHT = Hashtable.Create()
    Call myHT.Add("one", "The")
    Call myHT.Add("two", "quick")
    Call myHT.Add("three", "brown")
    Call myHT.Add("four", "fox")
    Call myHT.Add("five", "jumps")
    
    ' Displays the count and values of the Hashtable.
    Debug.Print "Initially,"
    Debug.Print Strings.Format("   Count    : {0}", myHT.Count)
    Debug.Print "   Values:"
    Call PrintKeysAndValues(myHT)
    
    ' Clears the Hashtable.
    Call myHT.Clear

    ' Displays the count and values of the Hashtable.
    Debug.Print "After Clear,"
    Debug.Print Strings.Format("   Count    : {0}", myHT.Count)
    Debug.Print "   Values:"
    Call PrintKeysAndValues(myHT)
End Sub

Private Sub PrintKeysAndValues(ByVal myHT As DotNetLib.Hashtable)
    Dim formatString As String
    formatString = Regex.Unescape("\t{0}:\t{1}")
    Debug.Print Regex.Unescape("\t-KEY-\t-VALUE-")
    Dim varDictEntry As Variant
    For Each varDictEntry In myHT
        Dim dictEntry As mscorlib.DictionaryEntry
        dictEntry = varDictEntry
        Debug.Print Strings.Format(formatString, DictionaryEntry.Key(dictEntry), DictionaryEntry.Value(dictEntry))
    Next
End Sub

'/*
'This code produces the following output.
'
'Initially,
'   Count    : 5
'Values:
'        -KEY-   -VALUE-
'two:            quick
'three:          brown
'four:           fox
'five:           jumps
'one:            The
'
'After Clear,
'   Count    : 0
'Values:
'        -KEY-   -VALUE-
'
'*/
