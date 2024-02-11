Attribute VB_Name = "ArrayFindAllPersonExample"
'@Folder("TestExamples.System.Array.Methods")

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 February 11, 2024
'@LastModified February 11, 2024

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.array.findall?view=netframework-4.8.1

'@Dependencies
'   Person.cls          Example data type used
'   PersonFilter.cls    Person filter implementing DotNetLib.IPredicate,
'                       Contains property CallBack returning a Predicate

Option Explicit

''
' Example of using a predicate to retrieve all the elements that match the conditions
' defined by the specified predicate.
' The conditions in the example are defined by a set of filters for a Person object as follows:
'   FirstName       Comparision ignores case and contains word
'   LastName        Comparision ignores case and contains word
'   DateOfBirth     Date of birth is greater or equal than specified
'   Age             Age is greater than or equal than specified
'   Status          Alive, Dead or AliveOrDead
'
' The predicate function PersonFilter.IsPersonMatch tests if a match is found according to
' the specified filters.
'
' @Remarks
'   The PersonFilter.cls implements a Predicate i.e. a function returning a boolean given
'   a parameter to match.
'   I.e. Private Function IPredicate_CallBack(ByVal pMatch As Variant) As Boolean
'
'   To create a Predicate assign the class implementing the Predicate using Predicate.Create
'   In the example the PersonFilter.CallBack returns the predicate Created using Predicate.Create(.Self)
'   Alternatively for example could create the Predicate object where MyPredicateClass.cls is a
'   predeclared class   implementing a Predicate as follows:
'     Dim myPredicate as DotNetLib.Predicate
'     Set myPredicate = Predicate.Create(MyPredicateClass)
'   Useage
'     Dim myResult as DotNetLib.Array
'     Set myResult = Arrays.FindAll(myArray, myPredicate)
'
'  Alternatively more predicates could have been used for each query on the array
'  to simplify the predicate logic.
''
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
                        Person.Create("Cameron", "Boyce", DateTime.CreateFromDate(1999, 5, 28), DateTime.CreateFromDate(2019, 5, 6)), _
                        Person.Create("John", "Harrismith", DateTime.CreateFromDate(1933, 7, 25)))
                        
    Dim pvtResult As DotNetLib.Array
    
    'Create a person predicate that contains a person with last name Smith  and greater and equal to 20 years old
    Dim pvtPersonFilter As PersonFilter
    Set pvtPersonFilter = PersonFilter.Create("", "Smith", Nothing, 20)

    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Debug.Print "People with the surname that contains " & pvtPersonFilter.LastName & " that are least " & pvtPersonFilter.Age & " years old."
    Call DisplayFindAll(pvtResult)
    Debug.Print
    
    pvtPersonFilter.Reset
    pvtPersonFilter.Age = 20
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Debug.Print "People that are at least of the age " & pvtPersonFilter.Age
    Call DisplayFindAll(pvtResult)
    
    Debug.Print
    pvtPersonFilter.Reset
    pvtPersonFilter.Status = PersonStatus.Deceased
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Debug.Print "Deceased people:"
    Call DisplayFindAll(pvtResult)
    
    Debug.Print
    pvtPersonFilter.Reset
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Debug.Print "Person List:"
    Call DisplayFindAll(pvtResult)
    
    'Testing no results
    Debug.Print
    pvtPersonFilter.Reset
    pvtPersonFilter.Age = 100
    Set pvtResult = Arrays.FindAll(personArray, pvtPersonFilter.CallBack)
    Debug.Print "People that are least of the age " & pvtPersonFilter.Age
    Call DisplayFindAll(pvtResult)
End Sub

Private Sub DisplayFindAll(ByVal personArray As DotNetLib.Array)
    If personArray Is Nothing Then
        Debug.Print "Person Array is nothing"
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
'    People with the surname that contains Smith that are least 20 years old.
'    Mary Smith (born January 1, 2000)
'    Bob Smith (born April 10, 1997)
'    Simon Smithfield (born August 17, 1950)
'    Martha Smithfield (born April 21, 1956)
'    John Harrismith (born July 25, 1933)
'
'    People that are at least of the age 20
'    Mary Smith (born January 1, 2000)
'    Mary Jones (born September 3, 1975)
'    Bob Smith (born April 10, 1997)
'    Simon Smithfield (born August 17, 1950)
'    Martha Smithfield (born April 21, 1956)
'    John Harrismith (born July 25, 1933)
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
'    John Harrismith (born July 25, 1933)
'
'    People that are least of the age 100
'    No results.


