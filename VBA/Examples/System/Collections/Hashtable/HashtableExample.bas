Attribute VB_Name = "HashtableExample"
'@Folder("Examples.System.Collections.Hashtable")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 19, 2023
'@LastModified October 19, 2023
'
'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb
'
'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1#examples

'@Remarks
' To enumerate ICollection Keys and ICollection Values cast to IEnumerable
' Bug Hashtable.Item(key) = valuetype causes an Object required error use SetValue(key,value) until fixed.
' To fix requires creating an IDL and manually adding a propput for value types and compiling type library with MIDL.

Option Explicit

''
' The following example shows how to create, initialize and perform various
' functions to a Hashtable and how to print out its keys and values.
''
Public Sub HashtableExample()
    Dim openWith As DotNetLib.Hashtable
    Set openWith = Hashtable.Create()
    
    ' Add some elements to the hash table. There are no
    ' duplicate keys, but some of the values are duplicates.
    Call openWith.Add("txt", "notepad.exe")
    Call openWith.Add("bmp", "paint.exe")
    Call openWith.Add("dib", "paint.exe")
    Call openWith.Add("rtf", "wordpad.exe")
    
    ' The Add method throws an exception if the new key is
    ' already in the hash table.
    On Error Resume Next
    Call openWith.Add("txt", "winword.exe")
    If Catch() Then
        Debug.Print "An element with Key = ""txt"" already exists."
    End If
    On Error GoTo 0 'Stop code and display error
    
    ' The Item property is the default property, so you
    ' can omit its name when accessing elements.
    Debug.Print Strings.Format("For key = ""rtf"", value = {0}.", openWith("rtf"))

    ' The default Item property can be used to change the value
    ' associated with a key.
    'Bug openWith.Item("rtf") = "winword.exe" causes an Object required error use SetValue instead until fixed.
    Call openWith.SetValue("rtf", "winword.exe")
    Debug.Print Strings.Format("For key = ""rtf"", value = {0}.", openWith("rtf"))
    
    ' If a key does not exist, setting the default Item property
    ' for that key adds a new key/value pair.
    Call openWith.SetValue("doc", "winword.exe")
    Debug.Print Strings.Format("For key = ""rtf"", value = {0}.", openWith("rtf"))
    
    ' ContainsKey can be used to test keys before inserting
    ' them.
     If (Not openWith.ContainsKey("ht")) Then
        Call openWith.Add("ht", "hypertrm.exe")
        Debug.Print Strings.Format("Value added for key = ""ht"": {0}", openWith("ht"))
     End If

    ' When you use foreach to enumerate hash table elements,
    ' the elements are retrieved as KeyValuePair objects.
    Debug.Print
    Dim varDictionaryEntry As Variant
    For Each varDictionaryEntry In openWith
        Dim dicEntry As DotNetLib.DictionaryEntry
        Set dicEntry = DictionaryEntry.Create(varDictionaryEntry)
        Debug.Print Strings.Format("Key = {0}, Value = {1}", dicEntry.Key, dicEntry.Value)
    Next

    ' To get the values alone, use the Values property.
    Dim valueColl As mscorlib.IEnumerable
    Set valueColl = openWith.Values

    ' The elements of the ValueCollection are strongly typed
    ' with the type that was specified for hash table values.
    Debug.Print
    Dim s As Variant
    For Each s In valueColl
        Debug.Print Strings.Format("Value = {0}", s)
    Next

    ' To get the keys alone, use the Keys property.
    Dim keyColl As mscorlib.IEnumerable
    Set keyColl = openWith.keys
    ' The elements of the KeyCollection are strongly typed
    ' with the type that was specified for hash table keys.
    Debug.Print
    Dim dicKey As Variant
    For Each dicKey In keyColl
        Debug.Print Strings.Format("Value = {0}", dicKey)
    Next
    
    ' Use the Remove method to remove a key/value pair.
    Debug.Print Regex.Unescape("\nRemove(""doc"")")
    Call openWith.Remove("doc")
    If (Not openWith.ContainsKey("doc")) Then
        Debug.Print "Key ""doc"" is not found."
    End If
End Sub

'/* This code example produces the following output:
'
'An element with Key = "txt" already exists.
'For key = "rtf", value = wordpad.exe.
'For key = "rtf", value = winword.exe.
'Value added for key = "ht": hypertrm.exe
'
'Key = dib, Value = paint.exe
'Key = txt, Value = notepad.exe
'Key = ht, Value = hypertrm.exe
'Key = bmp, Value = paint.exe
'Key = rtf, Value = winword.exe
'Key = doc, Value = winword.exe
'
'value = Paint.exe
'value = notepad.exe
'value = hypertrm.exe
'value = Paint.exe
'value = winword.exe
'value = winword.exe
'
'Key = dib
'Key = txt
'Key = ht
'Key = bmp
'Key = rtf
'Key = doc
'
'Remove ("doc")
'Key "doc" is not found.
' */
