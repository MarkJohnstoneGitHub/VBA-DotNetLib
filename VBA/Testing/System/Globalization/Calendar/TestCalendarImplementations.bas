Attribute VB_Name = "TestCalendarImplementations"
'@Folder("Testing.System.Globalization.Calendar")

Option Explicit

'@TODO: For Testing calendar implementations
'Lists Calendar types not currently implemented for default calendar or optional calendars
Public Sub CultureInfoCultureTypesCalendarTest()
    Dim cultures() As DotNetLib.CultureInfo
    cultures = CultureInfo.GetCultures(CultureTypes.CultureTypes_AllCultures)
    Dim varCultureInfo As Variant
    For Each varCultureInfo In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCultureInfo
        
        'Test if default calendar is implemented
        If culture.Calendar Is Nothing Then
            Debug.Print culture.EnglishName
        End If
        
        Dim varOptionalCalendar As Variant
        For Each varOptionalCalendar In culture.OptionalCalendars
            Dim pvtOptionalCalendar As DotNetLib.Calendar
            Set pvtOptionalCalendar = varOptionalCalendar
            'Test if optional calendar is implemented
            If pvtOptionalCalendar Is Nothing Then
                Debug.Print culture.EnglishName
            End If
        Next
        
    Next
End Sub

'Display optional calendars other  then GregorianCalendar
Public Sub CultureInfoCultureTypesCalendarTest2()
    Dim cultures() As DotNetLib.CultureInfo
    cultures = CultureInfo.GetCultures(CultureTypes.CultureTypes_AllCultures)
    Dim varCultureInfo As Variant
    For Each varCultureInfo In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCultureInfo
        
        'Test if default calendar is implemented
'        If culture.Calendar Is Nothing Then
'            Debug.Print culture.EnglishName
'        End If
        
        Dim varOptionalCalendar As Variant
        For Each varOptionalCalendar In culture.OptionalCalendars
            Dim pvtOptionalCalendar As DotNetLib.Calendar
            Set pvtOptionalCalendar = varOptionalCalendar
            'Test if optional calendar is implemented
            If Not (TypeOf pvtOptionalCalendar Is DotNetLib.GregorianCalendar) Then
                Debug.Print culture.EnglishName; "  "; pvtOptionalCalendar.ToString
            End If
        Next
        
    Next
End Sub


'List of Cultures using a calendar other than the Gregorian calendar
Public Sub CultureInfoCultureTypesCalendarNonGregorianCalendar()
    Dim cultures() As DotNetLib.CultureInfo
    cultures = CultureInfo.GetCultures(CultureTypes.CultureTypes_AllCultures)
    Dim varCultureInfo As Variant
    For Each varCultureInfo In cultures
        Dim culture As DotNetLib.CultureInfo
        Set culture = varCultureInfo
        If Not (TypeOf culture.Calendar Is DotNetLib.GregorianCalendar) Then
            Debug.Print culture.EnglishName; ", "; typeName(culture.Calendar)
        End If
    Next
End Sub

