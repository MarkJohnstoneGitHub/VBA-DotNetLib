Attribute VB_Name = "TestingArrayExamples"
'@Folder("Testing.System.Array")
Option Explicit

Private Sub TestingArrayCreateInstance()
'    Dim pType As DotNetLib.Type
'    Set pType = Types.GetType("DotNetLib.System.DateTime", False, False)
    'Set pType = Types.GetType("System.Text.StringBuilder", False, False)
        
    Dim doubleType As DotNetLib.Type
    
    Dim my1DArray As DotNetLib.Array
    Set my1DArray = Arrays.CreateInstance(Strings(), 5)
        
    Dim my1DArray2 As DotNetLib.Array
    Set my1DArray2 = Arrays.CreateInstance(Strings.GetType, 5)
    
'    Dim pType As DotNetLib.Type
'    Set pType = Types.GetType("System.DateTime")
    Dim myDateTimeArray As DotNetLib.Array
    Set myDateTimeArray = Arrays.CreateInstance(Types.GetType("System.DateTime"), 5)
    
    Dim myDateTimeArray2 As DotNetLib.Array
    Set myDateTimeArray2 = Arrays.CreateInstance(DateTime.GetType, 5)
    
    
    Dim personArray As DotNetLib.Array
    Set personArray = Arrays.CreateInstance(Person.GetType, 5)
    
    
End Sub

Private Sub TestingArraysCreateInstanceInt32()
    Dim longType As DotNetLib.Type
    Set longType = Objects.Create(CLng(0)).GetType()
    Dim longArray As DotNetLib.Array
    Set longArray = Arrays.CreateInstance(longType, 5)
    
    Dim longType2 As DotNetLib.Type
    Set longType2 = Types.GetType("System.Int32")
    Dim longArray2 As DotNetLib.Array
    Set longArray2 = Arrays.CreateInstance(longType2, 5)
  
    Dim longArray3 As DotNetLib.Array
    Set longArray3 = Arrays.CreateInstance(Int32.GetType, 5)
    
    Dim longArray4 As DotNetLib.Array
    Set longArray4 = Arrays.CreateInstance(Int32(), 5)
End Sub


Private Sub TestArrayDateTime()
    Dim myDateTime As DotNetLib.DateTime
    Set myDateTime = DateTime.CreateFromDate(2000, 1, 1)

    Dim datetimeArray As DotNetLib.Array
    Set datetimeArray = Arrays.CreateInstance(DateTime(), 5)
    
    datetimeArray.SetValue DateTime.CreateFromDate(2005, 6, 20).value, 0
    datetimeArray.SetValue Now(), 1
    datetimeArray.SetValue myDateTime.value, 2
    datetimeArray.SetValue DateTime.CreateFromDate(1983, 12, 31).value, 3
    datetimeArray.SetValue DateTime.CreateFromDate(1981, 10, 31).value, 4
    Arrays.Sort datetimeArray
    
    
End Sub

'Unsupported variant type
Private Sub TestArrayChar()
    Dim charArray As DotNetLib.Array
    Set charArray = Arrays.CreateInstance(Char.GetType, 5)
End Sub



'    Dim myType As DotNetLib.Type
'    Set myType = ObjectStatic.Create(CLng(0)).GetType()
'    Dim my1DArray As DotNetLib.Array
'    Set my1DArray = ArrayStatic.CreateInstance(myType, 5)
'    my1DArray.SetValue CLng(100), 0
'    Debug.Print my1DArray.Item(0)
    
'    Dim myType2 As DotNetLib.Type
'    Set myType2 = ObjectStatic.Create(Person.Create("", DateTime.CreateFromDate(2000, 1, 1))).GetType()
'    Dim my1DArrayV2 As DotNetLib.Array
'    Set my1DArrayV2 = ArrayStatic.CreateInstance(myType2, 5)
'    my1DArrayV2.SetValue Person.Create("Bobby", DateTime.CreateFromDate(2010, 5, 5)), 1
'
'    Dim per As Person
'    Set per = my1DArrayV2.Item(1)
'    Debug.Print per
'    Debug.Print my1DArrayV2.Item(1)
'    'InvalidCastException
'    my1DArrayV2.SetValue DateTime.CreateFromDate(2010, 5, 5), 2


