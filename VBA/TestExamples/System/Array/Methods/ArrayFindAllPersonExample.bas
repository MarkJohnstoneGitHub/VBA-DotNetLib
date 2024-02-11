Attribute VB_Name = "ArrayFindAllPersonExample"
'@Folder("TestExamples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 11, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.find?view=netframework-4.8.1

'@Dependencies
'   Person.cls          Example data type used
'   PersonFilter.cls    Person filter implementing DotNetLib.IPredicate,
'                       Contains property CallBack returning a Predicate

Option Explicit

Public Sub ArrayFindAllPersonExample()
    Dim personArray As DotNetLib.Array
    Set personArray = Arrays.CreateInitialize1D(Person.GetType, _
                        Person.Create("Mary", "Smith", DateTime.CreateFromDate(2000, 1, 1)), _
                        Person.Create("Mary", "Jones", DateTime.CreateFromDate(1975, 9, 3)), _
                        Person.Create("Bob", "Smith", DateTime.CreateFromDate(1997, 4, 10)), _
                        Person.Create("Luke", "Smith", DateTime.CreateFromDate(2020, 12, 3)), _
                        Person.Create("Simon", "Smithfield", DateTime.CreateFromDate(1950, 8, 17)), _
                        Person.Create("Martha", "Smithfield", DateTime.CreateFromDate(1956, 4, 21)), _
                        Person.Create("Zelda", "Sears", DateTime.CreateFromDate(1873, 1, 21), DateTime.CreateFromDate(1935, 2, 19)), _
                        Person.Create("Bob", "Hope", DateTime.CreateFromDate(1903, 5, 29), DateTime.CreateFromDate(2003, 7, 27)), _
                        Person.Create("Cameron", "Boyce", DateTime.CreateFromDate(1999, 5, 28), DateTime.CreateFromDate(2019, 5, 6)))
    
    Dim pvtResult As DotNetLib.Array
    
    'Create a person predicate that contains a person with last name Smith  and greater and equal to 20 years old
    Dim pvtPersonFilter As PersonFilter
    Set pvtPersonFilter = PersonFilter.Create("", "Smith", Nothing, 20)

'    Dim pvtPersonPredicate As DotNetLib.Predicate
'    Set pvtPersonPredicate = Predicate.Create(pvtPersonFilter)
'    Set pvtPersonPredicate = pvtPersonFilter

    Debug.Print "People with the surname that contains " & pvtPersonFilter.LastName & " that are least " & pvtPersonFilter.Age & " years old."
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Call DisplayFindAll(pvtResult)
    Debug.Print
    
    pvtPersonFilter.Reset
    pvtPersonFilter.Age = 22
    Debug.Print "People that are least of the age " & pvtPersonFilter.Age
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Call DisplayFindAll(pvtResult)
    
    Debug.Print
    pvtPersonFilter.Reset
    pvtPersonFilter.Status = PersonStatus.Deceased
    Debug.Print "Deceased people:"
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Call DisplayFindAll(pvtResult)
    
    Debug.Print
    pvtPersonFilter.Reset
    Debug.Print "Person List:"
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Call DisplayFindAll(pvtResult)
    
    'Testing no results
    Debug.Print
    pvtPersonFilter.Reset
    pvtPersonFilter.Age = 100
    Debug.Print "People that are least of the age " & pvtPersonFilter.Age
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Call DisplayFindAll(pvtResult)
End Sub

Private Sub DisplayFindAll(ByVal personArray As DotNetLib.Array)
    If personArray Is Nothing Then
        Exit Sub
    End If
    If personArray.Count = 0 Then
        Debug.Print "No results."
        Exit Sub
    End If
    
    Dim varItem As Variant
    For Each varItem In personArray
        Dim pvtPerson As Person
        Set pvtPerson = varItem
        Dim pvtDateOfDeath As String
        If pvtPerson.DateOfDeath Is Nothing Then
            pvtDateOfDeath = VBA.vbNullString
        Else
            pvtDateOfDeath = " – " & pvtPerson.DateOfDeath.ToString2("MMMM d, yyyy")
        End If
        
        Debug.Print VBString.Format("{0} {1} (born {2:MMMM d, yyyy}{3})", pvtPerson.FirstName, pvtPerson.LastName, _
                                    pvtPerson.DateOfBirth, _
                                    pvtDateOfDeath)
    Next
End Sub

' Output for example
'
'    People with the surname that contains smith that are least 20 years old.
'    Mary Smith (born January 1, 2000)
'    Bob Smith (born April 10, 1997)
'    Simon Smithfield (born August 17, 1950)
'    Martha Smithfield (born April 21, 1956)
'
'    People that are least of the age 22
'    Mary Smith (born January 1, 2000)
'    Mary Jones (born September 3, 1975)
'    Bob Smith (born April 10, 1997)
'    Simon Smithfield (born August 17, 1950)
'    Martha Smithfield (born April 21, 1956)
'
'    Deceased People:
'    Zelda Sears (born January 21, 1873 – February 19, 1935)
'    Bob Hope (born May 29, 1903 – July 27, 2003)
'    Cameron Boyce (born May 28, 1999 – May 6, 2019)
'
'    Person List:
'    Mary Smith (born January 1, 2000)
'    Mary Jones (born September 3, 1975)
'    Bob Smith (born April 10, 1997)
'    Luke Smith (born December 3, 2020)
'    Simon Smithfield (born August 17, 1950)
'    Martha Smithfield (born April 21, 1956)
'    Zelda Sears (born January 21, 1873 – February 19, 1935)
'    Bob Hope (born May 29, 1903 – July 27, 2003)
'    Cameron Boyce (born May 28, 1999 – May 6, 2019)
'
'    People that are least of the age 100
'    No results.

