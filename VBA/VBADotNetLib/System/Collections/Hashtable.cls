VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hashtable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Represents a collection of key/value pairs that are organized based on the hash code of the key."
'@Folder "VBADotNetLib.System.Collections"
'@PredeclaredId
'@ModuleDescription("Represents a collection of key/value pairs that are organized based on the hash code of the key.")

'@Singleton

'@Author Mark Johnstone
'@Project https://github.com/MarkJohnstoneGitHub/VBA-DotNetLib
'@Version v1.0 October 18, 2023
'@LastModified October 18, 2023

'@ReferenceAddin DotNetLib.tlb, mscorlib.tlb

'@Reference https://learn.microsoft.com/en-us/dotnet/api/system.collections.hashtable?view=netframework-4.8.1

'@Remarks
'Singleton class for Hashtable containing factory methods and static members.

Option Explicit

Private Type THashtableSingleton
   HashtableSingleton As DotNetLib.HashtableSingleton
End Type

Private this As THashtableSingleton

Private Sub Class_Initialize()
   Set this.HashtableSingleton = New DotNetLib.HashtableSingleton
End Sub

Private Sub Class_Terminate()
   Set this.HashtableSingleton = Nothing
End Sub

'@Description("Initializes a new, empty instance of the Hashtable class using the specified initial capacity and load factor, and the default hash code provider and comparer.")
Public Function Create(Optional ByVal capacity As Long = 0, Optional ByVal loadFactor As Single = 1, Optional ByRef equalityComparer As mscorlib.IEqualityComparer) As DotNetLib.Hashtable
Attribute Create.VB_Description = "Initializes a new, empty instance of the Hashtable class using the specified initial capacity and load factor, and the default hash code provider and comparer."
   Set Create = this.HashtableSingleton.Create(capacity, loadFactor, equalityComparer)
End Function

'@Description("Initializes a new instance of the Hashtable class by copying the elements from the specified dictionary to the new Hashtable object. The new Hashtable object has an initial capacity equal to the number of elements copied, and uses the specified load factor and IEqualityComparer object.")
Public Function Create2(ByVal d As mscorlib.IDictionary, Optional ByVal loadFactor As Single = 1, Optional ByVal equalityComparer As IEqualityComparer) As DotNetLib.Hashtable
Attribute Create2.VB_Description = "Initializes a new instance of the Hashtable class by copying the elements from the specified dictionary to the new Hashtable object. The new Hashtable object has an initial capacity equal to the number of elements copied, and uses the specified load factor and IEqualityComparer object."
   Set Create2 = this.HashtableSingleton.Create2(d, loadFactor, equalityComparer)
End Function
