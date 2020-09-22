VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form6 
   Caption         =   "Edit XML Document"
   ClientHeight    =   6780
   ClientLeft      =   2190
   ClientTop       =   1275
   ClientWidth     =   6795
   LinkTopic       =   "Form6"
   ScaleHeight     =   6780
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   2
      Left            =   5640
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   5520
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   1
      Left            =   3720
      Picture         =   "Form6.frx":08CA
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   5520
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Index           =   0
      Left            =   1680
      Picture         =   "Form6.frx":1194
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   5520
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":1A5E
            Key             =   "node_ico"
            Object.Tag             =   "node_ico"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":233A
            Key             =   "opened_node_ico"
            Object.Tag             =   "opened_node_ico"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":2C16
            Key             =   "text_ico"
            Object.Tag             =   "text_ico"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form6.frx":34F2
            Key             =   "root_ico"
            Object.Tag             =   "root_ico"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "XML Document Treeview"
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6495
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   6588
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CommandButton Command1 
         Caption         =   "Regenerate XML"
         Height          =   495
         Index           =   3
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear All"
         Height          =   495
         Index           =   2
         Left            =   2640
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View External DTD File"
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open a XML Document"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Right Click on Element to retrieve its attribute"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   6240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Text Element"
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Node Element"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Root Element"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Menu MnuPopUp 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAddRoot 
         Caption         =   "Add Root"
      End
      Begin VB.Menu MnuAddNode 
         Caption         =   "Add Child Node"
      End
      Begin VB.Menu MnuAddTextChild 
         Caption         =   "Add Text as Child"
      End
      Begin VB.Menu MnuAddTextNext 
         Caption         =   "Add Text as Next"
      End
      Begin VB.Menu MnuDelElement 
         Caption         =   "Delete Element"
      End
      Begin VB.Menu mnuATTMaint 
         Caption         =   "Retrieves Attribute"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xmlDoc As Msxml2.DOMDocument40
Dim cur_Date As Date, stop_Date As Date

Private Sub Command1_Click(Index As Integer)
On Error GoTo m_err:
Dim filePath As String, DocPath As String

If Index = 0 Then
    'clear all stuff first
    Command1.Item(2).Value = True
    
   'open a xml document
   ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|dtd Files (*.dtd)|*.dtd | XML Files (*.xml)|*.xml"
    ' Specify default filter
    CommonDialog1.FilterIndex = 3
    'set default path
    CommonDialog1.InitDir = App.Path
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    
    ' set file name to variable
    DocPath = CommonDialog1.filename
    
    'load xml document into treeview
    Call BuildTree(DocPath)
ElseIf Index = 1 Then
    If Not xmlDoc Is Nothing Then
      ' view external DTD
      If Not getDTDPath(filePath) Then
         GoTo m_err:
      Else
         Call Form7.execute(filePath)
      End If
    Else
      MsgBox "open a file first!", vbCritical
    End If
ElseIf Index = 2 Then     'clear all object
    TreeView1.Nodes.Clear
    Set xmlDoc = Nothing
    Set objXML = Nothing
ElseIf Index = 3 Then      'generate xml
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
    If Not getDTDPath(filePath) Then GoTo m_err:
    If Not filePath = "" Then ' check if DTD is selected
       Call lexh.startDTD(TreeView1.Nodes.Item(1), "SYSTEM", filePath)
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
       Call objFunction.traverseNode2XML2(TreeView1.Nodes.Item(1), cnth)
     End If
 
     Set objFunction = Nothing
     Set atrs = Nothing
     cnth.endElement "", "", TreeView1.Nodes.Item(1)  'end root element

    ' display xml to text box
    Call Form2.execute(wrt, filePath)
End If

m_quit:
  Exit Sub
m_err:
  'MsgBox "Error Occurs!", vbCritical
  GoTo m_quit:
End Sub

Private Sub BuildTree(filePath As String)
    On Error GoTo m_err:   'Error exception handler
    
    Dim i As Integer
    Dim TreeNode As Node                              ' node of treeview
    Dim currentRootElement As Msxml2.IXMLDOMElement   'Mother
    Dim currentNodeElement As Msxml2.IXMLDOMNode      'Child
    Dim tempRootElement As Msxml2.IXMLDOMElement
    Dim tempNodeMap As Msxml2.IXMLDOMNamedNodeMap
    Dim objFunction As New ClsFunction
    
    'Reset the Treeview
    Me.TreeView1.Nodes.Clear
    
    Set xmlDoc = New Msxml2.DOMDocument40
    xmlDoc.async = False
    If Not xmlDoc.Load(filePath) Then GoTo m_err:

    'The first element is always the Mother
    Set currentRootElement = xmlDoc.documentElement

    'add root to treeview
    If Not add2TreeWithImage(TreeView1, TreeNode, 0, 0, , currentRootElement.nodeName, 4) Then GoTo m_err: ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
    ' add attribute to objxml
    Call exploreAttribute(currentRootElement, TreeNode)
    
    'check if current mother has child
    If currentRootElement.hasChildNodes Then
       TreeNode.Expanded = True   'Expand the root to display its child
       Call objFunction.traverseNodeWithImage(TreeView1, currentRootElement, TreeNode)
    End If
    Set objFunction = Nothing

m_quit:
   Exit Sub
m_err:
   With xmlDoc.parseError
    Call displayError(.errorCode, .filePos, .line, .linePos, .reason, .srcText, .url)
  End With
  MsgBox "Can't open document!", vbInformation
  GoTo m_quit:
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set objXML = Nothing
End Sub


Private Sub MnuAddNode_Click()
Call AddNode
End Sub

Private Sub mnuAddRoot_Click()
Dim tempStr As String

'check if root exist
If checkRootExist Then
   MsgBox "Only one root for one XML document!", vbCritical
Else
   tempStr = InputBox("Enter name for Root Node(CASE Sensitive)")
   If Not add2TreeWithImage(TreeView1, , 0, 0, "", tempStr) Then GoTo m_err:
End If
   
m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub MnuAddTextChild_Click()
Call AddText("CHILD")
End Sub

Private Sub MnuAddTextNext_Click()
Call AddText("NEXT")
End Sub

Private Sub mnuATTMaint_Click()
Call ATTMaintenance
End Sub

Private Sub mnuDelElement_Click()
Call delElement
End Sub

Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   ' MsgBox "left click"
Else
   ' MsgBox "right click"
   PopupMenu MnuPopUp
End If
End Sub

Private Sub ATTMaintenance()
If Not TreeView1.SelectedItem.Tag = "TEXT" Then   'text node is not allow to have attributes
   Call Form5.execute(TreeView1.SelectedItem.Index)
End If
End Sub

Private Sub AddNode()
On Error GoTo m_err:
Dim tempIndex As Integer
Dim tempStr As String
Dim objAttribute As New ClsAttribute

 Set objAttribute.myNode = TreeView1.SelectedItem ' set reference
 
 tempIndex = TreeView1.SelectedItem.Index
 tempStr = InputBox("Enter name for the Child Node(CASE Sensitive)", , "New Node")
 If Not tempStr = "" Then 'check if tempStr is empty
    If Not add2TreeWithImage(TreeView1, , tempIndex, 4, "", tempStr, 1, 2) Then GoTo m_err: ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
    ' add a dummy objAttribute to objxml
    Call objXML.AddItem(myAttribute, objAttribute, tempIndex)
 Else
    MsgBox "Please enter a value!", vbCritical
 End If
 
Set objAttribute = Nothing
m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub AddText(tNavigation As String)
On Error GoTo m_err:
Dim tempIndex As Integer
Dim tempStr As String
Dim objAttribute As New ClsAttribute

 Set objAttribute.myNode = TreeView1.SelectedItem  'set reference
 
 tempIndex = TreeView1.SelectedItem.Index
 If Not tempIndex = 1 Then  'make sure the selected Index is not root node
 tempStr = InputBox("Enter name for Next Node(CASE Sensitive)", , "New Text")
 If Not tempStr = "" Then 'check if string is empty
    If tNavigation = "NEXT" Then
       If Not add2TreeWithImage(TreeView1, , tempIndex, 2, "", tempStr, 3) Then GoTo m_err: ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
    ElseIf tNavigation = "CHILD" Then
       If Not add2TreeWithImage(TreeView1, , tempIndex, 4, "", tempStr, 3) Then GoTo m_err: ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
    End If
    ' add a dummy objAttribute to objxml
    Call objXML.AddItem(myAttribute, objAttribute)
 Else
    MsgBox "Please enter a value!", vbCritical
 End If
Else
 MsgBox "Cannot add text to root node!", vbCritical
End If

Set objAttribute = Nothing
m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End Sub

Private Sub delElement()
 Dim TreeNode As Node
 Dim tempIndex As Integer
 Dim tempCol As New Collection
 Dim i As Integer
 
 
 Set TreeNode = TreeView1.SelectedItem
 If Not TreeNode Is Nothing Then
    Label2.Caption = "Deleting nodes .........PLEASE WAIT"  ' Process indicator
    tempIndex = TreeView1.SelectedItem.Index
    Call exploreChild(tempCol, TreeNode)       ' scan through the current node for all its child node's index
    Call removeATTORText(myAttribute, tempCol) ' remove attributes
    Call TreeView1.Nodes.Remove(tempIndex)     ' remove nodes from tree

    If Not checkRootExist Then
       MsgBox "WARNING!, root has been deleted!", vbCritical
    End If
 Else
    MsgBox "Select a Node first!", vbInformation
 End If

m_quit:
  Label2.Caption = ""       ' process indicator reset
  Set tempCol = Nothing
  Exit Sub
m_err:
  GoTo m_quit:
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
  MsgBox "Total nodes to be deleted = " & tCol.Count
  cur_Date = Now     'set currrent time
  While (counter <> tCol.Count)
    For i = 1 To tCol.Count - 1
       If tCol.Item(i).myValue < tCol.Item(i + 1).myValue Then
          tempIndex = tCol.Item(i).myValue
          tCol.Item(i).myValue = tCol.Item(i + 1).myValue
          tCol.Item(i + 1).myValue = tempIndex
       End If
       DoEvents
    Next i
    counter = counter + 1
  Wend
  Call calStopDateTime
End Sub

Private Sub calStopDateTime()
  On Error GoTo m_err:   'error exception handler
  Dim totalSec As Integer
  Dim tempVal As Double, fixVal As Integer, floatVal As Double, currentLevel As String
  Dim tempYears As Double, tempMonths As Double, tempDays As Double, tempHours As Double, tempMinutes As Double, tempSeconds As Double
  Dim myYears As Integer, myMonths As Integer, myDays As Integer, myHours As Integer, myMinutes As Integer, mySeconds As Integer
  Dim objFunction As ClsFunction
  stop_Date = Now
  
  totalSec = DateDiff("s", cur_Date, stop_Date)
  
  'convert second to year
  tempYears = CDbl(totalSec / 31104000)
  If tempYears > 0 Then
     tempVal = tempYears
     fixVal = Fix(tempYears)
     currentLevel = "YEAR"
  Else
     tempMonths = CDbl(totalSec / 2592000)
     If tempMonths > 0 Then
        tempVal = tempMonths
        fixVal = Fix(tempMonths)
        currentLevel = "MONTH"
     Else
       tempDays = CDbl(totalSec / 86400)
       If tempDays > 0 Then
          tempVal = tempDays
          fixVal = Fix(tempDays)
          currentLevel = "DAY"
       Else
          tempHours = CDbl(totalSec / (3600))
          If tempHours > 0 Then
             tempVal = tempHours
             fixVal = Fix(tempHours)
             currentLevel = "HOUR"
          Else
             tempMinutes = CDbl(totalSec / 60)
             If tempMinutes > 0 Then
                tempVal = tempMinutes
                fixVal = Fix(tempMinutes)
                currentLevel = "MINUTE"
             Else
                'seconds only
                GoTo m_sec_only:
             End If
          End If
       End If
     End If
  End If
  floatVal = Abs(tempVal - fixVal)
  Set objFunction = New ClsFunction
  myYears = 0
  myMonths = 0
  myDays = 0
  myHours = 0
  myMinutes = 0
  mySeconds = 0
  Call objFunction.calTimeDiff(floatVal, myYears, myMonths, myDays, myHours, myMinutes, mySeconds, currentLevel)
  Set objFunction = Nothing
  
  MsgBox "Time Taken = " & vbCrLf & _
        " Years = " & myYears & vbCrLf & _
        " Months = " & myMonths & vbCrLf & _
        " Days = " & myDays & vbCrLf & _
        " Hours = " & myHours & vbCrLf & _
        " Minutes = " & myMinutes & vbCrLf & _
        " Seconds = " & mySeconds, vbInformation
        
m_quit:
  Exit Sub
m_sec_only:
  MsgBox "Total Time Taken = " & totalSec, vbInformation
  GoTo m_quit:
m_err:
  MsgBox Err.Description & " YOU MAY IGNORE THIS ERROR!", vbInformation
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

Private Sub addRoot()
  Dim tempStr As String
  tempStr = InputBox("Enter name for Next Node(CASE Sensitive)", , "New Root")
  If Not tempStr = "" Then
     If Not add2TreeWithImage(TreeView1, , , , "", tempStr, 4) Then GoTo m_err:
  End If

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub
Private Function checkRootExist() As Boolean
If TreeView1.Nodes.Count > 0 Then
   checkRootExist = True
Else
   checkRootExist = False
End If

End Function

Private Function getDTDPath(tfilePath As String) As Boolean
'On Error GoTo m_err:
Dim dType As Msxml2.IXMLDOMDocumentType
Dim tempAttribute As IXMLDOMAttribute

getDTDPath = False
Set dType = xmlDoc.doctype
    If Not dType Is Nothing Then
       If xmlDoc.doctype.Attributes.length > 0 Then  ' check if there is any attributes
          Set tempAttribute = xmlDoc.doctype.Attributes.getQualifiedItem("SYSTEM", "")
          tfilePath = tempAttribute.nodeValue
       Else
          tfilePath = ""
       End If
    End If
getDTDPath = True

m_quit:
   Exit Function
m_err:
   GoTo m_quit:
End Function
