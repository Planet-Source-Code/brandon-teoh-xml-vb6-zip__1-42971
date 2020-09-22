VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form4 
   Caption         =   "Create XML Document"
   ClientHeight    =   8310
   ClientLeft      =   1380
   ClientTop       =   375
   ClientWidth     =   9585
   LinkTopic       =   "Form4"
   ScaleHeight     =   8310
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Create Schema Manually"
      Height          =   1335
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton Command1 
         Caption         =   "Create a simple XML Document"
         Height          =   735
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Generate XML"
      Height          =   495
      Left            =   3840
      TabIndex        =   13
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   7800
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Creating Schema Dynamically"
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   9615
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4095
         Left            =   2400
         TabIndex        =   20
         Top             =   2040
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7223
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Add Text to current node"
         Height          =   495
         Left            =   2400
         TabIndex        =   18
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Display Treeview Node to Debug-Window"
         Height          =   495
         Left            =   5400
         TabIndex        =   17
         Top             =   6240
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "ATTRIBUTE  MAINTENEANCE"
         Height          =   495
         Left            =   480
         TabIndex        =   12
         Top             =   6240
         Width           =   1815
      End
      Begin VB.TextBox txtDTD 
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   1320
         Width           =   6855
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Delete"
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   5040
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Clear All"
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Add Last"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   4440
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add Previous"
         Height          =   495
         Left            =   480
         TabIndex        =   6
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add Next"
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Select DTD File"
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Child"
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add Root Node"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "**** To Add, update, delete 'values' to a node, select a node and click 'Add Text' **********"
         Height          =   375
         Left            =   2400
         TabIndex        =   19
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "******To ADD, Update, Delete Attribute, click 'Attribute Maintenance ********"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "*************To Edit Node, Single Click on the Node *******"
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRTNode 
         Caption         =   "Add Root Node"
      End
      Begin VB.Menu mnuAddPrevious 
         Caption         =   "Add Previous"
      End
      Begin VB.Menu mnuAddNext 
         Caption         =   "Add Next"
      End
      Begin VB.Menu mnuAddChild 
         Caption         =   "Add Child"
      End
      Begin VB.Menu mnuAddLast 
         Caption         =   "Add Last"
      End
      Begin VB.Menu mnuATTMaint 
         Caption         =   "Attribute Maintenance"
      End
      Begin VB.Menu mnuTxtMaint 
         Caption         =   "Text Maintenance"
      End
      Begin VB.Menu mnuDelElement 
         Caption         =   "Delete Element"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myDTD As String

Private Sub Command1_Click()
 Dim rdr As New SAXXMLReader40
    Dim wrt As New MXXMLWriter40

    'We need these variables for typecasting the writer.
    Dim cnth As IVBSAXContentHandler
    Dim dtdh As IVBSAXDTDHandler
    Dim lexh As IVBSAXLexicalHandler
    Dim dech As IVBSAXDeclHandler
    Dim errh As IVBSAXErrorHandler
    
    'This is just a helper.
    Dim atrs As New SAXAttributes40
    
    'Set handler variables to the writer so it implements the interfaces.
    Set cnth = wrt
    Set dtdh = wrt
    Set lexh = wrt
    Set dech = wrt
    Set errh = wrt
    
    'set encoding because textbox object support UTF-8
    wrt.encoding = "UTF-8"
    ' using built-in DTD
    wrt.standalone = True
    
    ' Manually call necessary events to generate the XML file.
    cnth.startDocument
    lexh.startDTD "catalog", "", ""
        dech.elementDecl "catalog", "(book+)"
        dech.elementDecl "book", "(title | descr)"
        dech.attributeDecl "book", "author", "CDATA", "#IMPLIED", ""
        dech.attributeDecl "book", "ISBN", "CDATA", "#REQUIRED", _
          "000000000"
        dech.attributeDecl "book", "cover", "(hard|soft)", "", "soft"
        dech.elementDecl "title", "(#PCDATA)"
        dech.elementDecl "descr", "(#PCDATA)"
    lexh.endDTD
    cnth.startElement "", "", "catalog", atrs
    atrs.Clear
      atrs.addAttribute "", "", "ISBN", "", "0-06-097619-5"
      atrs.addAttribute "", "", "cover", "", "hard"
      cnth.startElement "", "", "book", atrs
        atrs.Clear
        cnth.startElement "", "", "title", atrs
        cnth.characters "On the Circular Problem of Quadratic Equations"
        cnth.endElement "", "", "title"
      cnth.endElement "", "", "book"
    cnth.endElement "", "", "catalog"
    
    Call Form2.execute(wrt, txtDTD.Text)
End Sub

Private Sub Command1_GotFocus()
Call enableOrDisableBut(False)
End Sub

Private Sub Command10_Click()
 Dim TreeNode As Node
 Dim tempIndex As Integer
 Dim tempCol As New Collection
 Dim i As Integer
 
 Set TreeNode = TreeView1.SelectedItem
 If Not TreeNode Is Nothing Then
    tempIndex = TreeView1.SelectedItem.Index
    Call exploreChild(tempCol, TreeNode)       ' scan through the current node for all its child node's index
    Call removeATTORText(myAttribute, tempCol) ' remove attributes
    Call removeATTORText(myText, tempCol)      ' remove text
    Call TreeView1.Nodes.Remove(tempIndex)     ' remove nodes from tree
    
    If Not checkRootExist Then
       Command2.Enabled = True
       MsgBox "WARNING!, root has been deleted!", vbCritical
    End If
Else
   MsgBox "Select a Node first!", vbInformation
End If

Call enableOrDisableBut(False)

m_quit:
  Set tempCol = Nothing
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub removeATTORText(myChoice As tSelect, tCol As Collection)
Dim i As Integer
For i = 1 To tCol.Count
    If myChoice = myAttribute Then
       'delete its attribute
       Call objXML.DelItem_Index(myAttribute, tCol.Item(i).myValue)
    Else
       'delete its text
       Call objXML.DelItem_Index(myText, tCol.Item(i).myValue)
    End If
Next i

End Sub
Private Sub exploreChild(ByRef tCol As Collection, SelectedNode As Node)
Dim i As Integer
Dim objFunction As ClsFunction
Dim objGeneral As New ClsGeneral

If SelectedNode.children > 0 Then
   Set objFunction = New ClsFunction  ' create new object so that the next process is create at a new stack, preventing stack overflow
   Call objFunction.scanChild(tCol, SelectedNode)
   objGeneral.myValue = SelectedNode.Index
   Call tCol.Add(objGeneral)  ' add object to collection
Else
   objGeneral.myValue = SelectedNode.Index
   Call tCol.Add(objGeneral)  ' add object to collection
End If
' rearrange all item in descending order
Call bubbleSort(tCol)

m_quit:
  Set objFunction = Nothing
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub bubbleSort(ByRef tCol As Collection)
'sort the item in ascending order
Dim i As Integer, counter As Integer
Dim tempIndex As Variant
counter = 1
  While (counter <> tCol.Count)
    For i = 1 To tCol.Count - 1
       If tCol.Item(i).myValue < tCol.Item(i + 1).myValue Then
          tempIndex = tCol.Item(i).myValue
          tCol.Item(i).myValue = tCol.Item(i + 1).myValue
          tCol.Item(i + 1).myValue = tempIndex
       End If
    Next i
    counter = counter + 1
  Wend
    
End Sub

Private Sub Command11_Click()
 
 Call Form5.execute(TreeView1.SelectedItem.Index)
 Call enableOrDisableBut(False)
End Sub

Private Sub Command12_Click()
Dim rdr As New SAXXMLReader40
    Dim wrt As New MXXMLWriter40

    'We need these variables for typecasting the writer.
    Dim cnth As IVBSAXContentHandler
    Dim dtdh As IVBSAXDTDHandler
    Dim lexh As IVBSAXLexicalHandler
    Dim dech As IVBSAXDeclHandler
    Dim errh As IVBSAXErrorHandler
    
    'SAX attribute class
    Dim atrs As New SAXAttributes40
    Dim i As Integer, j As Integer
    Dim objAttribute As ClsAttribute
    Dim objATTNameValue As ClsATTNameValue
    Dim objFunction As New ClsFunction
    
    'Set handler variables to the writer so it implements the interfaces.
    Set cnth = wrt
    Set dtdh = wrt
    Set lexh = wrt
    Set dech = wrt
    Set errh = wrt
   
    wrt.encoding = "UTF-8"  'set encoding because textbox object support UTF-8
    wrt.indent = True       'Configures the writer to indent elements.
   
    ' Manually call necessary events to generate the XML file.
    Call cnth.startDocument   'start document
    If Not txtDTD.Text = "" Then ' check if DTD is selected
       Call lexh.startDTD(TreeView1.Nodes.Item(1), "SYSTEM", txtDTD.Text)
       Call lexh.endDTD  'End DTD
    End If
    
           Set objAttribute = objXML.getItem_Index(myAttribute, 1)
           For j = 1 To objAttribute.getItemCount
               Set objATTNameValue = objAttribute.getItem_Index(j)
               atrs.addAttribute "", "", objATTNameValue.ATT_Name, "", objATTNameValue.ATT_Value
           Next j
     
     'add root and its corresponding attribute to to xml
     cnth.startElement "", "", TreeView1.Nodes.Item(1), atrs  'Start element tag
     atrs.Clear 'refresh SAX Attribute object
        
     'check if current mother has child
     If TreeView1.Nodes.Item(1).Expanded Then
       ' send to a method to perform loop, providing the root as parameter
       Call objFunction.traverseNode2XML(TreeView1.Nodes.Item(1), cnth)
     End If
 
     Set objFunction = Nothing
     Set atrs = Nothing
     cnth.endElement "", "", TreeView1.Nodes.Item(1)  'end root element

    'TextResult.Text = wrt.output
    Call Form2.execute(wrt, txtDTD.Text)

End Sub

Private Sub Command13_Click()
Dim objFunction As New ClsFunction

 'add root to treeview
 Debug.Print TreeView1.Nodes.Item(1)
        
 'check if current mother has child
 If TreeView1.Nodes.Item(1).Expanded Then
       ' send to a method to perform loop, providing the root as parameter
       Call objFunction.traverseNode2(TreeView1.Nodes.Item(1))
 End If
 
 Set objFunction = Nothing

End Sub

Private Sub Command14_Click()
Dim ObjText As ClsText

If Not TreeView1.SelectedItem.Index = 1 Then ' make sure not root node
      Set ObjText = objXML.getItem_Index(myText, TreeView1.SelectedItem.Index)
      ObjText.myData = InputBox("Enter/Update the text", "Element Maintenance", ObjText.myData)
Else
   MsgBox "Root Node cannot consist element!", vbCritical
End If

m_quit:
  Set ObjText = Nothing
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command2_Click()
Dim tempStr As String
Dim TreeNode As Node

tempStr = InputBox("Enter name for Root Node(CASE Sensitive)")
' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
If Not add2Tree(TreeView1, TreeNode, 0, 0, "", tempStr) Then
   GoTo m_err:
Else
   ' add a default attribute to objxml
   Dim objAttribute As New ClsAttribute
   'objAttribute.MyNode_Number = TreeNode.Index   ' set the index
   Set objAttribute.myNode = TreeNode                 ' set current node to objAttribute
   Call objXML.AddItem(myAttribute, objAttribute)
   
   'add a default text to objxml
   Dim ObjText As New ClsText 'create new element object
   ObjText.myData = ""
   Call objXML.AddItem(myText, ObjText)
    
   'destroy object
   Set objAttribute = Nothing
   Set ObjText = Nothing
End If

m_quit:
  If checkRootExist Then
     Command2.Enabled = False
     TreeNode.Expanded = True
  End If
  Exit Sub
m_err:
  MsgBox "Error occurs!", vbCritical
  GoTo m_quit:
End Sub

Private Sub Command3_Click()
Dim tempIndex As Integer
Dim tempStr As String
Dim TreeNode As Node

 tempIndex = TreeView1.SelectedItem.Index
 tempStr = InputBox("Enter name for the Child Node(CASE Sensitive)", , "New Node")
 If Not tempStr = "" Then 'check if tempStr is empty
    If Not add2Tree(TreeView1, TreeNode, tempIndex, 4, "", tempStr) Then ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
       GoTo m_err:
    Else
       ' add a default object to objxml
       Dim objAttribute As New ClsAttribute
       'objAttribute.MyNode_Number = TreeNode.Index   'set the index
       Set objAttribute.myNode = TreeNode
       Call objXML.AddItem(myAttribute, objAttribute)
       
       'add a default text to objxml
       Dim ObjText As New ClsText 'create new element object
       ObjText.myData = ""
       Call objXML.AddItem(myText, ObjText)
   
       Set objAttribute = Nothing
       Set ObjText = Nothing
    End If
 Else
    MsgBox "Please enter a value!", vbCritical
 End If
 
 Call enableOrDisableBut(False)

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command4_Click()
On Error GoTo m_err:
' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|dtd Files (*.dtd)|*.dtd"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    'set default path
    CommonDialog1.InitDir = App.Path
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    
    ' set file name to variable
    myDTD = CommonDialog1.filename
    
    txtDTD.Text = myDTD

    If Not txtDTD.Text = "" Then Command4.Caption = "Change DTD File"

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Dim tempIndex As Integer
Dim tempStr As String
Dim TreeNode As Node

tempIndex = TreeView1.SelectedItem.Index
If Not tempIndex = 1 Then  'make sure the selected Index is not root node
 tempStr = InputBox("Enter name for Next Node(CASE Sensitive)", , "New Node")
 If Not tempStr = "" Then 'check if string is empty
    If Not add2Tree(TreeView1, TreeNode, tempIndex, 2, "", tempStr) Then ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
       GoTo m_err:
    Else
       ' add a default object to objxml
       Dim objAttribute As New ClsAttribute
       'objAttribute.MyNode_Number = TreeNode.Index   'set the index
       Set objAttribute.myNode = TreeNode
       Call objXML.AddItem(myAttribute, objAttribute)
       
       'add a default text to objxml
       Dim ObjText As New ClsText 'create new element object
       ObjText.myData = ""
       Call objXML.AddItem(myText, ObjText)
   
       Set objAttribute = Nothing
       Set ObjText = Nothing
    End If
 Else
    MsgBox "Please enter a value!", vbCritical
 End If
Else
 MsgBox "Cannot add next node to root node!", vbCritical
End If

Call enableOrDisableBut(False)

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command7_Click()
Dim tempIndex As Integer
Dim tempStr As String
Dim TreeNode As Node

tempIndex = TreeView1.SelectedItem.Index
If Not tempIndex = 1 Then  'make sure the selected Index is not root node
 tempStr = InputBox("Enter name for Previous Node(CASE Sensitive)", , "New Node")
 If Not tempStr = "" Then 'check if string is empty
    If Not add2Tree(TreeView1, TreeNode, tempIndex, 3, "", tempStr) Then ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
       GoTo m_err:
    Else
       ' add a default object to objxml
       Dim objAttribute As New ClsAttribute
       'objAttribute.MyNode_Number = TreeNode.Index   'set the index
       Set objAttribute.myNode = TreeNode
       Call objXML.AddItem(myAttribute, objAttribute)
       
       'add a default text to objxml
       Dim ObjText As New ClsText 'create new element object
       ObjText.myData = ""
       Call objXML.AddItem(myText, ObjText)
   
       Set objAttribute = Nothing
       Set ObjText = Nothing
    End If
 Else
    MsgBox "Please enter a value!", vbCritical
 End If
Else
 MsgBox "Cannot add previous to root node!", vbCritical
End If

Call enableOrDisableBut(False)

m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End Sub

Private Sub Command8_Click()
Dim tempIndex As Integer
Dim tempStr As String
Dim TreeNode As Node

tempIndex = TreeView1.SelectedItem.Index
If Not tempIndex = 1 Then  'make sure the selected Index is not root node
 tempStr = InputBox("Enter name for Last Node(CASE Sensitive)", , "New Node")
 If Not tempStr = "" Then 'check if tempStr is empty
    If Not add2Tree(TreeView1, TreeNode, tempIndex, 1, "", tempStr) Then ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
       GoTo m_err:
    Else
       ' add a default object to objxml
       Dim objAttribute As New ClsAttribute
       'objAttribute.MyNode_Number = TreeNode.Index   'set the index
       Set objAttribute.myNode = TreeNode
       Call objXML.AddItem(myAttribute, objAttribute)
       
       'add a default text to objxml
       Dim ObjText As New ClsText 'create new element object
       ObjText.myData = ""
       Call objXML.AddItem(myText, ObjText)
       
       Set objAttribute = Nothing
       Set ObjText = Nothing
    End If
 Else
    MsgBox "Please enter a value!", vbCritical
 End If
Else
 MsgBox "Cannot add last to root node!", vbCritical
End If

Call enableOrDisableBut(False)

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command9_Click()
TreeView1.Nodes.Clear
Command2.Enabled = True
Call enableOrDisableBut(False)
Set objXML = Nothing
End Sub

Private Function checkRootExist() As Boolean
If TreeView1.Nodes.Count > 0 Then
   checkRootExist = True
Else
   checkRootExist = False
End If

End Function

Private Sub Form_Load()
Call enableOrDisableBut(False)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set objXML = Nothing
End Sub

Private Sub mnuAddChild_Click()
If Command3.Enabled Then
   Command3.Value = True
Else
   MsgBox "Please select a node!", vbInformation
End If
End Sub

Private Sub mnuAddLast_Click()
If Command8.Enabled Then
   Command8.Value = True
Else
   MsgBox "Please select a node!", vbInformation
End If
End Sub

Private Sub mnuAddNext_Click()
If Command6.Enabled Then
   Command6.Value = True
Else
   MsgBox "Please select a node!", vbInformation
End If
End Sub

Private Sub mnuAddPrevious_Click()
If Command7.Enabled Then
   Command7.Value = True
Else
   MsgBox "Please select a node!", vbInformation
End If
End Sub

Private Sub mnuAddRTNode_Click()
   Command2.Value = True
End Sub


Private Sub TextSource_GotFocus()
 Call enableOrDisableBut(False)
End Sub

Private Sub mnuATTMaint_Click()
If Command11.Enabled Then
   Command11.Value = True
Else
   MsgBox "Please select a node!", vbCritical
End If
End Sub

Private Sub mnuDelElement_Click()
If Command10.Enabled Then
  Command10.Value = True
Else
   MsgBox "Please select a node first!", vbCritical
End If
End Sub

Private Sub mnuTxtMaint_Click()
If Command14.Enabled Then
   Command14.Value = True
Else
   MsgBox "Please select a node!", vbCritical
End If
End Sub


Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   ' MsgBox "left click"
Else
   ' MsgBox "right click"
   PopupMenu mnuPopUp
End If
End Sub

Private Sub enableOrDisableBut(b As Boolean)
Command7.Enabled = b
Command6.Enabled = b
Command3.Enabled = b
Command8.Enabled = b
Command11.Enabled = b
Command14.Enabled = b
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

Call enableOrDisableBut(True)
End Sub

Private Sub txtDTD_GotFocus()
 Call enableOrDisableBut(False)
End Sub
