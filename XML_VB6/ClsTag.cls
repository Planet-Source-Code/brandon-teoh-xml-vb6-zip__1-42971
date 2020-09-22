VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ATT_Col As Collection
Private cur_NodeIndex As Integer

Public Sub execute()
Set ATT_Col = New Collection
End Sub

Public Function getItemCount() As Integer
 getItemCount = ATT_Col.Count
End Function

Public Sub AddItem(objAttribute As ClsAttribute, tNodeIndex As Integer)
Call ATT_Col.Add(objAttribute, tNodeIndex)
End Sub

Public Function getItem(tNodeIndex As Integer) As ClsAttribute
getItem = ATT_Col.Item(CStr(tNodeIndex))
End Function

Public Sub DelItem(tNodeIndex As Integer)
ATT_Col.Remove (CStr(tNodeIndex))
End Sub

Public Sub setNodeIndex(index As Integer)
cur_NodeIndex = index
End Sub
