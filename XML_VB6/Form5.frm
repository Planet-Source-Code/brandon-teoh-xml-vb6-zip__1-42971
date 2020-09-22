VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form5 
   Caption         =   "ATTRIBUTE MAINTENANCE"
   ClientHeight    =   3900
   ClientLeft      =   1875
   ClientTop       =   2625
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   3900
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Attribute"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update Attribute"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Attribute"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Value :"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Attribute Name :"
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objAttribute As ClsAttribute
Dim objATTNameValue As ClsATTNameValue
Dim cur_NodeIndex As Integer
Dim tempKey As String

Private Sub Command1_Click(Index As Integer)
Dim tempCol As New Collection
Dim i As Integer, iIndex As Integer

If Index = 0 Then
  If Len(Text1(0).Text) <> 0 And Len(Text1(1).Text) <> 0 Then ' data validation
    'tempKey = objAttribute.MyNode_Number & "_" & Text1(0).Text & "_" & Text1(1).Text
    'tempKey = Text1(0).Text & "_" & Text1(1).Text
    tempKey = Text1(0).Text
    If Not objAttribute.isItemExist(tempKey, iIndex) Then 'check if already exist
      ' new ATT_Name_Value object
      Set objATTNameValue = New ClsATTNameValue
      objATTNameValue.ATT_Name = Text1(0).Text
      objATTNameValue.ATT_Value = Text1(1).Text
      'provide the key
      objATTNameValue.myKey = tempKey
 
      For i = 0 To Text1.UBound
       tempCol.Add Text1(i).Text
      Next i

      'add to listview
      Call add2listview(ListView1, CStr(objATTNameValue.myKey), tempCol)
      'add to objAttribute collection
      Call objAttribute.AddItem(objATTNameValue, objATTNameValue.myKey) 'objXml would be updated automatically
      'check if current item doesn't exist in objXML
      'If Not isItemExist Then Call add2ClsXML
    Else
      MsgBox "Attribute already exist, goto update!", vbInformation
    End If
  Else
    MsgBox "Please Enter Data Properly!", vbInformation
  End If
ElseIf Index = 1 Then ' update attribute
   
   For i = 0 To Text1.UBound
       tempCol.Add Text1(i).Text
   Next i
      
   'update 2 listview
   Call edit2listView(ListView1, tempCol)
   
   ' update the new values to objAttNamevalue
   objATTNameValue.ATT_Value = Text1(1).Text
   
   'objxml will be updated automatically
   
   Call resetButton("AFTER_UPDATE")
Else ' delete attributes
  'delete from listview
  Call deleteFromListView(ListView1)
  
  ' delete from objAttribute
  Call objAttribute.DelItem_key(tempKey)
  
  'objxml will be updated automatically
  
  'reset button status
  Call resetButton("AFTER_DELETE")
End If

m_quit:
  Text1(0).Text = ""
  Text1(1).Text = ""
  Set tempCol = Nothing
  Set objATTNameValue = Nothing
  tempKey = ""
  Exit Sub
m_err:
  GoTo m_quit:
End Sub
Private Sub deleteFromListView(lv As ListView)
lv.ListItems.Remove (CStr(tempKey))

End Sub

Private Sub edit2listView(ByRef lv As ListView, col As Collection)

Dim i As Integer

For i = 1 To col.Count
 With lv
      .SelectedItem.SubItems(i) = CStr(col.Item(i))
 End With
Next

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub InslvHeader()
 With ListView1
  .ColumnHeaders.Add , , "Key", 5    '0
  .ColumnHeaders.Add , , "ATT-Name", 2000   '1
  .ColumnHeaders.Add , , "ATT-Value", 2000 '2
 End With

End Sub

Public Sub execute(NodeIndex As Integer)
 cur_NodeIndex = NodeIndex
 
 ' Initialize listview header
 Call InslvHeader
 'load existing attributes
 Call loadAttributes

 Me.Show vbModal
End Sub

Private Sub add2listview(ByRef lv As ListView, tKey As String, col As Collection)
'on error goto m_err:

Dim i As Integer

Set ItemRetn = lv.ListItems.Add(, tKey, tKey)

For i = 1 To col.Count
    ItemRetn.SubItems(i) = col.Item(i) & ""
Next
                                      
m_quit:
  Set ItemRetn = Nothing
  Exit Sub
m_err:
  MsgBox "Error at add2listview", vbCritical
  GoTo m_quit:
  
End Sub

Private Sub loadAttributes()
 Dim objATTNameValue As ClsATTNameValue
 Dim i As Integer
 'If objXML.isItemExist(myAttribute, cur_NodeIndex) Then 'check if attribute exist
     
    'Set objAttribute = objXML.getItem_key(myAttribute, cur_NodeIndex) 'Create object reference here
    Set objAttribute = objXML.getItem_Index(myAttribute, cur_NodeIndex) 'create object reference here
    
    If objAttribute.getItemCount <> 0 Then
       For i = 1 To objAttribute.getItemCount
         Dim tempCol As New Collection
         Set objATTNameValue = objAttribute.getItem_Index(i)
             Call tempCol.Add(objATTNameValue.ATT_Name)
             Call tempCol.Add(objATTNameValue.ATT_Value)
             Call add2listview(ListView1, objATTNameValue.myKey, tempCol)
         Set tempCol = Nothing
       Next
    End If
 'Else
    'create objAttribute
    'Set objAttribute = New ClsAttribute
    'objAttribute.MyNode_Number = cur_NodeIndex
 'End If

m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End Sub

Private Sub Form_Load()
Command1(1).Enabled = False
Command1(2).Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'check if objAttribute has object in its collection

Set objAttribute = Nothing
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
   
   ' get the key from listview
   tempKey = ListView1.SelectedItem.Key
   
   ' retrieve objATTNamevalue from objattribute
   Set objATTNameValue = objAttribute.getItem_key(tempKey)
   ' publish values to the form
   Text1(0).Text = objATTNameValue.ATT_Name
   Text1(0).Enabled = False
   Text1(1).Text = objATTNameValue.ATT_Value
   Text1(1).SetFocus
   
   Command1(1).Enabled = True
   Command1(2).Enabled = True
   Command1(0).Enabled = False
   
End Sub

Private Sub resetButton(str As String)
If str = "AFTER_UPDATE" Then
  Text1(0).Enabled = True
  Command1(0).Enabled = True
  Command1(1).Enabled = False
  Command1(2).Enabled = False
ElseIf str = "AFTER_DELETE" Then
  Text1(0).Enabled = True
  Command1(0).Enabled = True
  Command1(1).Enabled = False
  Command1(2).Enabled = False
End If
End Sub
