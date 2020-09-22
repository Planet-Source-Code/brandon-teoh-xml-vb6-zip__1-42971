VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim ATT_Col As New Collection
Dim ATT_Key_Col As New Collection
Dim Text_Col As New Collection
Dim Text_Key_Col As New Collection

Public Function getItemCount(tempSelect As tSelect) As Integer
 If tempSelect = myAttribute Then
    getItemCount = ATT_Col.Count
 Else
    getItemCount = Text_Col.Count
 End If
End Function

Public Sub AddItem(tempSelect As tSelect, objTemp As Object, Optional tKey As Integer)
 If tempSelect = myAttribute Then
    If tKey <> 0 Then
       'Add to Item Collection
       Call ATT_Col.Add(objTemp, CStr(tKey))
       'Add to Key Collection
       Call ATT_Key_Col.Add(tKey)
    Else
       'add to item collection without key
       Call ATT_Col.Add(objTemp)
    End If
 Else
    If tKey <> 0 Then
       'Add to Item Collection
       Call Text_Col.Add(objTemp, CStr(tKey))
       'Add to Key Collection
       Call Text_Key_Col.Add(tKey)
    Else
       'add to item collection without the key
       Call Text_Col.Add(objTemp)
    End If
 End If
End Sub

Public Function getItem_key(tempSelect As tSelect, tNodeIndex As Integer) As Object
 If tempSelect = myAttribute Then
    Set getItem_key = ATT_Col.Item(CStr(tNodeIndex))
 Else
    Set getItem_key = Text_Col.Item(CStr(tNodeIndex))
 End If
End Function

Public Function getItem_Index(tempSelect As tSelect, i As Integer) As Object
 If tempSelect = myAttribute Then
    Set getItem_Index = ATT_Col.Item(i)
 Else
    Set getItem_Index = Text_Col.Item(i)
 End If
End Function
Public Sub DelItem_key(tempSelect As tSelect, tNodeIndex As Integer)
  If tempSelect = myAttribute Then
     'remove from item collection
     Call ATT_Col.Remove(CStr(tNodeIndex))
     'remove from key collection
     Call RemoveFromKeyCol(myAttribute, tNodeIndex)
   Else
     ' remove from item collection
     Text_Col.Remove (CStr(tNodeIndex))
     ' remove from key collection
     Call RemoveFromKeyCol(myText, tNodeIndex)
   End If
End Sub

Public Sub DelItem_Index(tempSelect As tSelect, i As Integer, Optional withKey As String)
 If tempSelect = myAttribute Then
    'remove from item collection
    ATT_Col.Remove (i)
    If withKey <> "" Then
       'remove from key collection
       ATT_Key_Col.Remove (i)
    End If
 Else
    ' remove from item collection
    Text_Col.Remove (i)
    If withKey <> "" Then
       'remove from key collection
       Text_Key_Col.Remove (i)
    End If
 End If
End Sub

Private Sub RemoveFromKeyCol(tempSelect As tSelect, tKey As Variant)
 Dim i As Integer
 If tempSelect = myAttribute Then
    'search through key collection
    For i = 1 To ATT_Key_Col.Count
      If ATT_Key_Col.Item(i) = tKey Then
         'remove from it
         Call ATT_Key_Col.Remove(i)
        Exit Sub
      End If
    Next
 Else
   'search through key collection
    For i = 1 To Text_Key_Col.Count
      If Text_Key_Col.Item(i) = tKey Then
         'remove from it
         Call Text_Key_Col.Remove(i)
         Exit Sub
      End If
    Next
 End If
End Sub

