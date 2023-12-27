Attribute VB_Name = "RegionInfoCreateExample2"
'@Folder("Examples.System.Globalization.RegionInfo.Constructors")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 December 27, 2023
'@LastModified December 27, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.globalization.regioninfo.-ctor?view=netframework-4.8.1#system-globalization-regioninfo-ctor(system-string)

Option Explicit

''
' The following code example creates instances of RegionInfo using culture names.
''
Public Sub RegionInfoCreateExample2()
    ' Creates an array containing culture names.
    Dim myCultures() As String
    Call ArrayEx.CreateInitialize1D(myCultures, "", "ar", "ar-DZ", "en", "en-US")

    ' Creates a RegionInfo for each of the culture names.
    '    Note that "ar" is the culture name for the neutral culture "Arabic",
    '    but it is also the region name for the country/region "Argentina";
    '    therefore, it does not fail as expected.
    Debug.Print ("Without checks...")
    Dim varCulture As Variant
    Dim myRI As DotNetLib.RegionInfo
    For Each varCulture In myCultures
        On Error Resume Next
        Set myRI = RegionInfo.Create2(varCulture)
        If Err.Number = ArgumentException Then
            Debug.Print Err.Description
        End If
        On Error GoTo 0 'Stop code and display error
    Next
    
    Debug.Print

    Debug.Print "Checking the culture names first..."
    For Each varCulture In myCultures
        If (varCulture = "") Then
            Debug.Print "The culture is the invariant culture."
        Else
            Dim myCI As DotNetLib.CultureInfo
            Set myCI = CultureInfo.CreateFromName(varCulture, False)
            If (myCI.IsNeutralCulture) Then
                Debug.Print VBAString.Format("The culture {0} is a neutral culture.", varCulture)
            Else
                Debug.Print VBAString.Format("The culture {0} is a specific culture.", varCulture)
                On Error Resume Next
                Set myRI = RegionInfo.Create2(varCulture)
                If Err.Number = ArgumentException Then
                    Debug.Print Err.Description
                End If
                On Error GoTo 0 'Stop code and display error
            End If
        End If
    Next
End Sub

'/*
'This code produces the following output.
'
'Without checks...
'There is no region associated with the Invariant Culture (Culture ID: 0x7F).
'The region name en should not correspond to neutral culture; a specific culture name is required.
'Parameter name: name
'
'Checking the culture names first...
'The culture is the invariant culture.
'The culture ar is a neutral culture.
'The culture ar-DZ is a specific culture.
'The culture en is a neutral culture.
'The culture en-US is a specific culture.
'
'*/
