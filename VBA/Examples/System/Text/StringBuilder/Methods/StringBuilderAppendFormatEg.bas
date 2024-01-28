Attribute VB_Name = "StringBuilderAppendFormatEg"
'@Folder("Examples.System.Text.StringBuilder.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 January 28, 2023
'@LastModified January 28, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.text.stringbuilder.appendformat?view=netframework-4.8.1#system-text-stringbuilder-appendformat(system-string-system-object)

Option Explicit

Private sb As DotNetLib.StringBuilder

''
' The following example demonstrates the AppendFormat method.
''
Public Sub StringBuilderAppendFormatExample()
    Set sb = StringBuilder.Create()
    
    Dim var1 As Long
    var1 = 111
    Dim var2 As Single
    var2 = 2.22
    Dim var3 As String
    var3 = "abcd"
    Dim var4() As Variant
    Call VBArray.CreateInitialize1D(var4, 3, 4.4, "X")
    
    Debug.Print
    Debug.Print "StringBuilder.AppendFormat method:"
    Call sb.AppendFormat("1) {0}", var1)
    Call Show(sb)
    Call sb.AppendFormat2("2) {0}, {1}", var1, var2)
    Call Show(sb)
    Call sb.AppendFormat3("3) {0}, {1}, {2}", var1, var2, var3)
    Call Show(sb)
    Call sb.AppendFormat4("4) {0}, {1}, {2}", var4)
    Call Show(sb)
    Dim ci As DotNetLib.CultureInfo
    Set ci = CultureInfo.CreateFromName("es-ES", True)
    Call sb.AppendFormat5(ci, "5) {0}", var2)
    Call Show(sb)
End Sub

Private Sub Show(ByVal sbs As DotNetLib.StringBuilder)
    Debug.Print sbs.ToString()
    sb.Length = 0
End Sub

'/*
'This example produces the following results:
'
'StringBuilder.AppendFormat Method:
'1) 111
'2) 111, 2.22
'3) 111, 2.22, abcd
'4) 3, 4.4, X
'5) 2,22
'*/

'using System;
'using System.Text;
'using System.Globalization;
'
'Class Sample
'{
'    static StringBuilder sb = new StringBuilder();
'
'    public static void Main()
'    {
'    int    var1   = 111;
'    float  var2   = 2.22F;
'    string var3   = "abcd";
'    object[] var4 = {3, 4.4, 'X'};
'
'    Console.WriteLine();
'    Console.WriteLine("StringBuilder.AppendFormat method:");
'    sb.AppendFormat("1) {0}", var1);
'    Show(sb);
'    sb.AppendFormat("2) {0}, {1}", var1, var2);
'    Show(sb);
'    sb.AppendFormat("3) {0}, {1}, {2}", var1, var2, var3);
'    Show(sb);
'    sb.AppendFormat("4) {0}, {1}, {2}", var4);
'    Show(sb);
'    CultureInfo ci = new CultureInfo("es-ES", true);
'    sb.AppendFormat(ci, "5) {0}", var2);
'    Show(sb);
'    }
'
'    public static void Show(StringBuilder sbs)
'    {
'    Console.WriteLine(sbs.ToString());
'    sb.Length = 0;
'    }
'}
'/*
'This example produces the following results:
'
'StringBuilder.AppendFormat Method:
'1) 111
'2) 111, 2.22
'3) 111, 2.22, abcd
'4) 3, 4.4, X
'5) 2,22
'*/
