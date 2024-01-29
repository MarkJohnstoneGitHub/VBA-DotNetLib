Attribute VB_Name = "StringLastIndexOfExample8"
'@Folder "Examples.System.Strings.Methods"

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 3, 2024
'@LastModified January 3, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.lastindexof?view=netframework-4.8.1#system-string-lastindexof(system-string)

Option Explicit

''
' The following example removes opening and closing HTML tags from a string if
' the tags begin and end the string. If a string ends with a closing bracket
' character (">"), the example uses the LastIndexOf method to locate the start
' of the end tag.
''
Public Sub StringLastIndexOfExample8()
    Dim strSource() As DotNetLib.String
    Call VBArray.CreateInitialize1D(strSource, Strings.Create("<b>This is bold text</b>"), Strings.Create("<H1>This is large Text</H1>"), _
                Strings.Create("<b><i><font color=green>This has multiple tags</font></i></b>"), _
                Strings.Create("<b>This has <i>embedded</i> tags.</b>"), _
                Strings.Create("This line ends with a greater than symbol and should not be modified>"))
                
    ' Strip HTML start and end tags from each string if they are present.
    Dim s As Variant
    For Each s In strSource
        Debug.Print "Before: " + s.ToString
        Dim Item As DotNetLib.String
        Set Item = s
        ' Use EndsWith to find a tag at the end of the line.
        If (Item.Trim().EndsWith3(">")) Then
            ' Locate the opening tag.
            Dim endTagStartPosition As Long
            endTagStartPosition = Item.LastIndexOf7("</")
            ' Remove the identified section, if it is valid.
            If (endTagStartPosition >= 0) Then
                Set Item = Item.Substring2(0, endTagStartPosition)
            End If

            ' Use StartsWith to find the opening tag.
            If (Item.Trim().StartsWith3("<")) Then
                ' Locate the end of opening tab.
                Dim openTagEndPosition As Long
                openTagEndPosition = Item.IndexOf7(">")
                ' Remove the identified section, if it is valid.
                If (openTagEndPosition >= 0) Then
                    Set Item = Item.Substring(openTagEndPosition + 1)
                End If
            End If
        End If
        ' Display the trimmed string.
        Debug.Print "After: " + Item.ToString()
        Debug.Print
    Next
End Sub

' The example displays the following output:
'    Before: <b>This is bold text</b>
'    After: This is bold text
'
'    Before: <H1>This is large Text</H1>
'    After: This is large Text
'
'    Before: <b><i><font color=green>This has multiple tags</font></i></b>
'    After: <i><font color=green>This has multiple tags</font></i>
'
'    Before: <b>This has <i>embedded</i> tags.</b>
'    After: This has <i>embedded</i> tags.
'
'    Before: This line ends with a greater than symbol and should not be modified>
'    After: This line ends with a greater than symbol and should not be modified>


