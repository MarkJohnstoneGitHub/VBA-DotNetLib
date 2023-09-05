Attribute VB_Name = "TextInfoCloneExample"
'@Folder("Examples.System.Globalization.TextInfo.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 September 5, 2023
'@LastModified September 5, 2023

'@Reference
' https://learn.microsoft.com/en-us/dotnet/api/system.globalization.textinfo.clone?view=netframework-4.8.1#examples
Option Explicit

' The following code example demonstrates the Clone and ReadOnly methods.
Public Sub TextInfoClone()
    ' Get the TextInfo of a predefined culture that ships with
    ' the .NET Framework.
    Dim ci As DotNetLib.CultureInfo
    Set ci = CultureInfo.CreateFromName("en-US")
    Dim ti1 As DotNetLib.TextInfo
    Set ti1 = ci.TextInfo
    
    ' Display whether the TextInfo is read-only or not.
    DisplayReadOnly "1) The original TextInfo object", ti1
    Debug.Print
    
    ' Create a clone of the original TextInfo and cast the clone to a TextInfo type.
    Debug.Print "2a) Create a clone of the original TextInfo object..."
    Dim ti2 As DotNetLib.TextInfo
    Set ti2 = ti1.Clone()
    
    ' Display whether the clone is read-only.
    DisplayReadOnly "2b) The TextInfo clone", ti2
    
    ' Set the ListSeparator property on the TextInfo clone.
    Debug.Print "2c) The original value of the clone's ListSeparator " + _
                "property is "; """"; ti2.ListSeparator; """"
    ti2.ListSeparator = "/"
    Debug.Print "2d) The new value of the clone's ListSeparator " + _
                "property is "; """"; ti2.ListSeparator; """"; VBA.vbNewLine
    
    ' Create a read-only clone of the original TextInfo.
    Debug.Print "3a) Create a read-only clone of the original TextInfo object..."
    Dim ti3 As DotNetLib.TextInfo
    Set ti3 = TextInfo.ReadOnly(ti1)
    ' Display whether the read-only clone is actually read-only.
    DisplayReadOnly "3b) The TextInfo clone", ti3
    
    ' Try to set the ListSeparator property of a read-only TextInfo object. Use the
    ' IsReadOnly property again to determine whether to attempt the set operation. You
    ' could use a try-catch block instead and catch an InvalidOperationException when
    ' the set operation fails, but that programming technique is inefficient.
    Debug.Print "3c) Try to set the read-only clone's LineSeparator " + _
                "property."
    If (ti3.IsReadOnly = True) Then
        Debug.Print "3d) The set operation is invalid."
    Else
        ' This clause is not executed.
        ti3.ListSeparator = "/"
        Debug.Print "3d) The new value of the clone's ListSeparator " + _
                    "property is "; """"; ti3.ListSeparator; """"; VBA.vbNewLine
    End If
End Sub

Private Sub DisplayReadOnly(ByVal caption As String, ByVal ti As DotNetLib.TextInfo)
    Debug.Print caption; " is "; IIf(ti.IsReadOnly, VBA.vbNullString, "not "); "read-only."
End Sub

'/*
'This code example produces the following results:
'
'1) The original TextInfo object is not read-only.
'
'2a) Create a clone of the original TextInfo object...
'2b) The TextInfo clone is not read-only.
'2c) The original value of the clone's ListSeparator property is ",".
'2d) The new value of the clone's ListSeparator property is "/".
'
'3a) Create a read-only clone of the original TextInfo object...
'3b) The TextInfo clone is read-only.
'3c) Try to set the read-only clone's LineSeparator property.
'3d) The set operation is invalid.
'
'*/

