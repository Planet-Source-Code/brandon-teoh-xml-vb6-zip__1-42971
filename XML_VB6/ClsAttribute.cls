VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public MyNode_Number As Integer
Public myNode As Node
Public ATTNameVal_Col As New Collection
Public ATTNameVal_Key_Col As New Collection

Public Function getItemCount() As Integer
 getItemCount = ATTNameVal_Col.Count
End Function

Public Sub AddItem(objATTNameVal As ClsATTNameValue, tKey As String)
 'Add to Item Collection
 Call ATTNameVal_Col.Add(objATTNameVal, tKey)
 'Add to key collection
 Call ATTNameVal_Key_Col.Add(tKey)
End Sub

Public Function getItem_key(tKey As String) As ClsATTNameValue
 Set getItem_key = ATTNameVal_Col.Item(CStr(tKey))
End Function

Public Function getItem_Index(i As Integer) As ClsATTNameValue
 Set getItem_Index = ATTNameVal_Col.Item(i)
End Function
Public Sub DelItem_key(tKey As String)
 Dim tempIndex As Integer
 
 ' remove from item collection
 Call ATTNameVal_Col.Remove(CStr(tKey))
 
 ' remove from key collection
 If isItemExist(tKey, tempIndex) Then
    Call ATTNameVal_Key_Col.Remove(tempIndex)
 End If
 
End Sub

Public Function isItemExist(tKey As Variant, ByRef Index As Integer) As Boolean
 Dim i As Integer, underScore_Pos1 As Integer, underScore_Pos2 As Integer
 isItemExist = False
 'check from Key collection
 For i = 1 To ATTNameVal_Key_Col.Count
     If ATTNameVal_Key_Col.Item(i) = tKey Then
          Index = i   'return current index as reference
          isItemExist = True
          Exit Function
     End If
 Next i
End Function

