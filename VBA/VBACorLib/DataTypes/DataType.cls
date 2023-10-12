VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBACorLib.DataTypes")
'@PredeclaredId

'@Todo Move each data type to its associated singleton class
'@Remarks Tempory Data type class until move data types to associated singleton class.

Option Explicit

Const DateTimeType          As String = "DotNetLib.System.DateTime"
Const DateTimeOffsetType    As String = "DotNetLib.System.DateTimeOffset"
Const Int16Type             As String = "System.Int16"
Const Int32Type             As String = "System.Int32"
Const Int64Type             As String = "System.Int64"
Const StringType            As String = "System.String"
Const TimeSpanType          As String = "DotNetLib.System.TimeSpan"

Private Type TDataType
    DateTime    As DotNetLib.Type
    DateTimeOffset  As DotNetLib.Type
    Int16           As DotNetLib.Type
    Int32Type       As DotNetLib.Type
    Int64Type       As DotNetLib.Type
    StringType      As DotNetLib.Type
    TimeSpanType    As DotNetLib.Type
End Type

Private this As TDataType

Private Sub Class_Initialize()
    Set this.DateTime = Types.GetType(DateTimeType)
    Set this.DateTimeOffset = Types.GetType(DateTimeOffsetType)
    Set this.Int16 = Types.GetType(Int16Type)
    Set this.Int32Type = Types.GetType(Int32Type)
    Set this.Int64Type = Types.GetType(Int64Type)
    Set this.StringType = Types.GetType(StringType)
    Set this.TimeSpanType = Types.GetType(TimeSpanType)
End Sub

Private Sub Class_Terminate()
    Set this.DateTime = Nothing
End Sub

Public Function GetDateTime() As DotNetLib.Type
    Set GetDateTime = this.DateTime
End Function

Public Function GetInt16() As DotNetLib.Type
    Set GetInt16 = this.Int16
End Function

Public Function GetInt32() As DotNetLib.Type
    Set GetInt32 = this.Int32Type
End Function

Public Function GetInt64() As DotNetLib.Type
    Set GetInt64 = this.Int64Type
End Function

Public Function GetString() As DotNetLib.Type
    Set GetString = this.StringType
End Function

Public Function GetTimeSpan() As DotNetLib.Type
    Set GetTimeSpan = this.TimeSpanType
End Function
