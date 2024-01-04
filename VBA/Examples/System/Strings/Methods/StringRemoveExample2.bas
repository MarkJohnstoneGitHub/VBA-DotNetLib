Attribute VB_Name = "StringRemoveExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 4, 2024
'@LastModified January 4, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.remove?view=netframework-4.8.1#system-string-remove(system-int32-system-int32)

Option Explicit

''
' The following example demonstrates how you can remove the middle name from a
' complete name.
''
Public Sub StringRemoveExample2()
    Dim pvtName As DotNetLib.String
    Set pvtName = Strings.Create("Michelle Violet Banks")
    
    Debug.Print VBAString.Format("The entire name is '{0}'", pvtName)
    
    ' Remove the middle name, identified by finding the spaces in the name.
    Dim founds1 As Long
    founds1 = pvtName.IndexOf7(" ")
    Dim foundS2 As Long
    foundS2 = pvtName.IndexOf9(" ", founds1 + 1)
    
    If (founds1 <> foundS2 And founds1 >= 0) Then
        Set pvtName = pvtName.Remove2(founds1 + 1, foundS2 - founds1)

        Debug.Print VBAString.Format("After removing the middle name, we are left with '{0}'", pvtName)
    End If
End Sub

' The example displays the following output:
'       The entire name is 'Michelle Violet Banks'
'       After removing the middle name, we are left with 'Michelle Banks'
