VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDiskKey"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function DiskKeyValidate(KeyDrive As String) As Boolean
    DiskKeyValidate = ValidateDisk(KeyDrive)
End Function

Public Function DiskKeyCreate(KeyName As String, KeyNum As String, KeyDrive As String) As Boolean
    DiskKeyCreate = CreateDiskKey(KeyName, KeyNum, KeyDrive)
End Function

Public Function DiskKeyName(KeyDrive As String) As String
    DiskKeyName = GetName(KeyDrive)
End Function

Public Function DiskKeyNum(KeyDrive As String) As String
    DiskKeyNum = GetKey(KeyDrive)
End Function
