Attribute VB_Name = "UriSegmentsExample"
'@Folder("Examples.System.Uri.Properties")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.segments?view=netframework-4.8.1

'@Remarks
'   VBString.Format4(format,arg0, arg1, arg2) used instead of VBString.Format
'   due to VBA compiler issue with ParamArray for this example
'   Alternative solution assign uriAddress1.Segments to a String Arrary. For example the following works as
'   expected.
'   Dim pvtSegments() as String
'   pvtSegments = uriAddress1.Segments
'   Debug.Print VBString.Format("The parts are {0}, {1}, {2}", pvtSegments(0), pvtSegments(1), pvtSegments(2))
'
'   Alternatively casting to a string works as expected using VBString.Format() eg.
'   VBString.Format("The parts are {0}, {1}, {2}", CStr(uriAddress1.Segments()(0)), CStr(uriAddress1.Segments()(1)), CStr(uriAddress1.Segments()(2)))

'
'@Issues Using VBString.Format(ByRef pFormat As String, ParamArray args() As Variant) As String
'@Reference https://stackoverflow.com/questions/3375562/vba-what-is-causing-this-string-argument-passed-to-paramarray-to-get-changed-to

Option Explicit

''
' The following example creates a Uri instance with 3 segments and displays
' the segments on the screen.
''
Public Sub UriSegmentsExample()
    Dim uriAddress1 As DotNetLib.Uri
    Set uriAddress1 = Uri.Create("http://www.contoso.com/title/index.htm")
    Debug.Print VBString.Format4("The parts are {0}, {1}, {2}", uriAddress1.Segments()(0), uriAddress1.Segments()(1), uriAddress1.Segments()(2))
End Sub

'Output
'    The parts are /, title/, index.htm
