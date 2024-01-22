Attribute VB_Name = "UriSegmentsExample2"
'@Folder("Examples.System.Uri.Properties")
'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 23, 2023
'@LastModified January 23, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.uri.segments?view=netframework-4.8.1#remarks

Option Explicit

''
' The following example shows the absolute path and segments for two URIs.
' The second example illustrates that the fragment and query are not part of
' the absolute path and therefore are not segments.
''
Public Sub UriSegmentsExample2()
    Dim uriAddress1 As DotNetLib.Uri
    Set uriAddress1 = Uri.Create("http://www.contoso.com/Chapters/Chapter1/Sections/Section1.htm")
    Dim pvtSegments() As String
    pvtSegments = uriAddress1.Segments
    
    Dim pvtIndex As Long
    For pvtIndex = LBound(pvtSegments) To UBound(pvtSegments)
        If pvtIndex = LBound(pvtSegments) Then
             Debug.Print VBString.Format2("The parts are {0}", pvtSegments(0));
        Else
            Debug.Print VBString.Format2(", {0}", pvtSegments(pvtIndex));
        End If
    Next
    Debug.Print
    
    Set uriAddress1 = Uri.Create("http://www.contoso.com/Chapters/Chapter1/Sections/Section1.htm#page1?answer=NO")
    pvtSegments = uriAddress1.Segments
    For pvtIndex = LBound(pvtSegments) To UBound(pvtSegments)
        If pvtIndex = LBound(pvtSegments) Then
             Debug.Print VBString.Format2("The parts are {0}", pvtSegments(0));
        Else
            Debug.Print VBString.Format2(", {0}", pvtSegments(pvtIndex));
        End If
    Next
End Sub

' Output
'    The parts are /, Chapters/, Chapter1/, Sections/, Section1.htm
'    The parts are /, Chapters/, Chapter1/, Sections/, Section1.htm


