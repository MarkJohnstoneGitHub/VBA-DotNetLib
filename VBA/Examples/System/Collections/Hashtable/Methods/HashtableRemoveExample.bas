Attribute VB_Name = "HashtableRemoveExample"
'@Folder("Examples.System.Collections.Hashtable.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 22, 2023
'@LastModified October 22, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable.remove?view=netframework-4.8.1#examples

Option Explicit

''
' The following example shows how to remove elements from the Hashtable.
''
Public Sub HashtableRemove()
    ' Creates and initializes a new Hashtable.
    Dim myHT As DotNetLib.Hashtable
    Set myHT = Hashtable.Create()
    Call myHT.Add("1a", "The")
    Call myHT.Add("1b", "quick")
    Call myHT.Add("1c", "brown")
    Call myHT.Add("2a", "fox")
    Call myHT.Add("2b", "jumps")
    Call myHT.Add("2c", "over")
    Call myHT.Add("3a", "the")
    Call myHT.Add("3b", "lazy")
    Call myHT.Add("3c", "dog")
    
    '/ Displays the Hashtable.
    Debug.Print "The Hashtable initially contains the following:"
    Call PrintKeysAndValues(myHT)
    
    ' Removes the element with the key "3b".
    Call myHT.Remove("3b")
    
    ' Displays the current state of the Hashtable.
    Debug.Print "After removing ""lazy"":"
    Call PrintKeysAndValues(myHT)
End Sub

Private Sub PrintKeysAndValues(ByVal myHT As DotNetLib.Hashtable)
    Dim varDictEntry As Variant
    For Each varDictEntry In myHT
        Dim dictEntry As DotNetLib.DictionaryEntry
        Set dictEntry = DictionaryEntry.Create(varDictEntry)
        Debug.Print Strings.Format("    {0}:    {1}", dictEntry.Key, dictEntry.Value)
    Next
    Debug.Print
End Sub

'/*
'This code produces the following output.
'
'The Hashtable initially contains the following:
'2 c:       over
'3 a:       The
'2 b:       jumps
'3 b:       lazy
'1 b:       quick
'3 c:       dog
'2 a:       fox
'1 c:       brown
'1 a:       The
'
'After removing "lazy":
'2 c:       over
'3 a:       The
'2 b:       jumps
'1 b:       quick
'3 c:       dog
'2 a:       fox
'1 c:       brown
'1 a:       The
'
'*/
