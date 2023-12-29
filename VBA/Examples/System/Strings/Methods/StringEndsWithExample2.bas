Attribute VB_Name = "StringEndsWithExample2"
'@Folder("Examples.System.Strings.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 29, 2023
'@LastModified December 29, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.string.endswith?view=netframework-4.8.1

Option Explicit

''
' The following example defines a StripEndTags method that uses the EndsWith(String)
' method to remove HTML end tags from the end of a line. Note that the StripEndTags
' method is called recursively to ensure that multiple HTML end tags at the end of
' the line are removed.
''
Public Sub StringEndsWithExample2()
    ' process an input file that contains html tags.
    ' this sample checks for multiple tags at the end of the line, rather than simply
    ' removing the last one.
    ' note: HTML markup tags always end in a greater than symbol (>).
                                    
    Dim strSource() As DotNetLib.String
    Call ArrayEx.CreateInitialize1D(strSource, Strings.Create("<b>This is bold text</b>"), _
                                Strings.Create("<H1>This is large Text</H1>"), _
                                Strings.Create("<b><i><font color=green>This has multiple tags</font></i></b>"), _
                                Strings.Create("<b>This has <i>embedded</i> tags.</b>"), _
                                Strings.Create("This line simply ends with a greater than symbol, it should not be modified>"))
                                    
                                    
    Debug.Print "The following lists the items before the ends have been stripped:"
    Debug.Print "-----------------------------------------------------------------"
    
    ' print out the initial array of strings
    Dim s As Variant
    For Each s In strSource
        Debug.Print s
    Next

    Debug.Print

    Debug.Print "The following lists the items after the ends have been stripped:"
    Debug.Print "----------------------------------------------------------------"

    ' print out the array of strings
    For Each s In strSource
        Debug.Print StripEndTags(s)
    Next
End Sub

Private Function StripEndTags(ByVal item As DotNetLib.String) As DotNetLib.String
    Dim found As Boolean
    found = False

    ' try to find a tag at the end of the line using EndsWith
    If (item.Trim().EndsWith3(">")) Then

        ' now search for the opening tag...
        Dim lastLocation As Long
        lastLocation = item.LastIndexOf7("</")

        ' remove the identified section, if it is a valid region
        If (lastLocation >= 0) Then
            found = True
            Set item = item.Substring2(0, lastLocation)
        End If
    End If

    If (found) Then
        Set item = StripEndTags(item)
    End If
    Set StripEndTags = item
End Function

' The example displays the following output:
'    The following lists the items before the ends have been stripped:
'    -----------------------------------------------------------------
'    <b>This is bold text</b>
'    <H1>This is large Text</H1>
'    <b><i><font color=green>This has multiple tags</font></i></b>
'    <b>This has <i>embedded</i> tags.</b>
'    This line simply ends with a greater than symbol, it should not be modified>
'
'    The following lists the items after the ends have been stripped:
'    ----------------------------------------------------------------
'    <b>This is bold text
'    <H1>This is large Text
'    <b><i><font color=green>This has multiple tags
'    <b>This has <i>embedded</i> tags.
'    This line simply ends with a greater than symbol, it should not be modified>
