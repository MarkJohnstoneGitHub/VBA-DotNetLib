Attribute VB_Name = "StringBuilderInsertExample"
'@Folder("Examples.System.Text.StringBuilder.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 31, 2023
'@LastModified October 31, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.insert?view=netframework-4.8.1#examples

Option Explicit

Private sb As DotNetLib.StringBuilder
Private initialValue As String

''
' The following example demonstrates the Insert method.
''
Public Sub StringBuilderInsert()
    initialValue = "--[]--"
    
    Dim xyz As String
    xyz = "xyz"
'    Dim obj As Object
'    Set obj = Nothing

    Dim xBool As Boolean
    xBool = True
    Dim xByte As Byte
    xByte = 1
    Dim xInt16 As Integer
    xInt16 = 2
    Dim xInt32 As Long
    xInt32 = 3
    Dim xInt64 As LongLong
    xInt64 = 4
    Dim xSingle As Single
    xSingle = 6.6
    Dim xDouble As Double
    xDouble = 7.7
    
    Debug.Print "StringBuilder.Insert method"
    Set sb = StringBuilder.Create(initialValue)
    
    Call sb.Insert(3, xyz)
    Call show(1, sb)
    
    Call sb.Insert_2(3, xyz, 2)
    Call show(2, sb)
    
    Call sb.Insert_3(3, xBool)      ' True
    Call show(3, sb)
    
    Call sb.Insert_4(3, xByte)      ' 1
    Call show(4, sb)
    
    Call sb.Insert_5(3, xInt16)     ' 2
    Call show(5, sb)
    
    Call sb.Insert_6(3, xInt32)     ' 3
    Call show(6, sb)

    Call sb.Insert_7(3, xInt64)     ' 4
    Call show(7, sb)
    
    Call sb.Insert_8(3, xSingle)      ' 6.6
    Call show(8, sb)

    Call sb.Insert_9(3, xDouble)      ' 7.7
    Call show(9, sb)
    
End Sub

Private Sub show(ByVal overloadNumber As Long, ByVal sbs As DotNetLib.StringBuilder)
    Debug.Print VBAString.Format("{0,2:G} = {1}", overloadNumber, sbs.ToString())
    Set sb = StringBuilder.Create(initialValue)
End Sub

'/*
'This example produces the following results:
'
'StringBuilder.Insert method
' 1 = --[xyz]--
' 2 = --[xyzxyz]--
' 3 = --[True]--
' 4 = --[1]--
' 5 = --[2]--
' 6 = --[3]--
' 7 = --[4]--
' 8 = --[6.6]--
' 9 = --[7.7]--
'*/
