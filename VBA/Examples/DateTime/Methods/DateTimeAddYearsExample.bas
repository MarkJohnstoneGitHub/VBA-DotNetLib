Attribute VB_Name = "DateTimeAddYearsExample"
'@Folder("VBADotNetLib.Examples.DateTime.Methods")

'https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 July 10, 2023
'@LastModified July 30, 2023

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.datetime.addyears?view=netframework-4.8.1

Option Explicit

' The following example illustrates using the AddYears method with a DateTime value that represents a leap year day.
' It displays the date for the fifteen years prior to and the fifteen years that follow February 29, 2000.
Public Sub DateTimeAddYears()
   Dim baseDate As IDateTime
   Set baseDate = DateTime.CreateFromDate(2000, 2, 29)
   Debug.Print "    Base Date:        " & baseDate.ToString2("d") & VBA.vbNewLine

   ' Show dates of previous fifteen years.
   Dim ctr As Long
   For ctr = -1 To -15 Step -1
      Debug.Print Abs(ctr) & " year(s) ago:        " & baseDate.AddYears(ctr).ToString2("d")
   Next
   Debug.Print VBA.vbNewLine

   ' Show dates of next fifteen years.
   For ctr = 1 To 15
      Debug.Print ctr & " year(s) ago:        " & baseDate.AddYears(ctr).ToString2("d")
   Next
End Sub

' The example displays the following output:
'           Base Date:        2/29/2000
'
'        1 year(s) ago:        2/28/1999
'        2 year(s) ago:        2/28/1998
'        3 year(s) ago:        2/28/1997
'        4 year(s) ago:        2/29/1996
'        5 year(s) ago:        2/28/1995
'        6 year(s) ago:        2/28/1994
'        7 year(s) ago:        2/28/1993
'        8 year(s) ago:        2/29/1992
'        9 year(s) ago:        2/28/1991
'       10 year(s) ago:        2/28/1990
'       11 year(s) ago:        2/28/1989
'       12 year(s) ago:        2/29/1988
'       13 year(s) ago:        2/28/1987
'       14 year(s) ago:        2/28/1986
'       15 year(s) ago:        2/28/1985
'
'        1 year(s) from now:   2/28/2001
'        2 year(s) from now:   2/28/2002
'        3 year(s) from now:   2/28/2003
'        4 year(s) from now:   2/29/2004
'        5 year(s) from now:   2/28/2005
'        6 year(s) from now:   2/28/2006
'        7 year(s) from now:   2/28/2007
'        8 year(s) from now:   2/29/2008
'        9 year(s) from now:   2/28/2009
'       10 year(s) from now:   2/28/2010
'       11 year(s) from now:   2/28/2011
'       12 year(s) from now:   2/29/2012
'       13 year(s) from now:   2/28/2013
'       14 year(s) from now:   2/28/2014
'       15 year(s) from now:   2/28/2015
