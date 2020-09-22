VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Learning XML using MSXML4"
   ClientHeight    =   5940
   ClientLeft      =   660
   ClientTop       =   1275
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Get Attributes of Element"
      Height          =   1695
      Left            =   6600
      TabIndex        =   16
      Top             =   4080
      Width           =   2175
      Begin VB.CommandButton Command6 
         Caption         =   "Method 2"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Method 1"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   6600
      TabIndex        =   11
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton Command5 
         Caption         =   "Get All Text of element"
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Get Elements of a xml Doc"
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Get Doctype of a xml Doc"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Get First-Child of a xml Doc"
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Understanding the difference between 'Child', 'Doctype', 'Elements' and 'text' of a xml document"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   2040
      TabIndex        =   8
      Top             =   0
      Width           =   4455
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3975
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   7011
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Clear All"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   4320
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Open a XML Document"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Close"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Parsing Document with inconsistent DTD"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Parsing Document with external DTD"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Parsing Document with internal DTD"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Display In TreeView"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Parse URL"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parse XMLDOMDocument"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'create the reader
Dim rdr As New SAXXMLReader40
'create the writer
Dim wrt As New MXXMLWriter40
Dim fileURL As String, xmlDoc As Msxml2.DOMDocument40

On Error GoTo errorHandler

wrt.byteOrderMark = True
wrt.omitXMLDeclaration = False
wrt.indent = True

'set the writer to the content handler
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt

'get File reference
If Not getURL(fileURL) Then GoTo m_err:

'parse the XML
rdr.parseURL fileURL

'show the results in a message box, which can only display a maximum of around 1019 characters
MsgBox wrt.output

m_quit:
    Exit Sub
errorHandler:
    MsgBox Err.Description
    GoTo m_quit:
m_err:
    MsgBox "Error occurs!", vbCritical
    GoTo m_quit:
End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command11_Click()
On Error GoTo m_err:
'create the reader
Dim rdr As New SAXXMLReader40
'create the writer
Dim wrt As New MXXMLWriter40
Dim fileURL As String, xmlDoc As Msxml2.DOMDocument40

wrt.byteOrderMark = True
wrt.omitXMLDeclaration = False
wrt.indent = True

'set the writer to the content handler
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt

 ' Set CancelError is True
    CommonDialog1.CancelError = True
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|xml Files" & _
    "(*.xml)|*.xml|dtd Files (*.dtd)|*.dtd"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file
    
    'MsgBox CommonDialog1.FileName
    
    If Not getDoc(xmlDoc, CommonDialog1.filename) Then GoTo m_err:

    'parse the XML
    rdr.parse xmlDoc

    'show the results in a message box, which can only display a maximum of around 1019 characters
    MsgBox wrt.output
    
m_quit:
    Exit Sub
m_err:
    GoTo m_quit:
End Sub

Private Sub Command12_Click()
TreeView1.Nodes.Clear
End Sub

Private Sub Command13_Click()
Dim xmlDoc As Msxml2.DOMDocument40
Dim objAttribute As IXMLDOMAttribute
Dim NoAtrs As Integer, AtrsName As String, AtrsVal As Variant
Dim i As Integer

'get File reference
If Not getDoc(xmlDoc, App.Path & "\XML_Dtd\db.xml") Then GoTo m_err:
MsgBox "No.of Attribute = " & xmlDoc.firstChild.Attributes.length
For i = 0 To xmlDoc.firstChild.Attributes.length - 1
  Set objAttribute = xmlDoc.firstChild.Attributes.Item(i)
  MsgBox "Attribute Name=" & objAttribute.nodeName & vbCrLf & _
         "Attribute Value=" & objAttribute.nodeValue
Next i
m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End Sub

Private Sub Command14_Click()
Dim xmlDoc As Msxml2.DOMDocument40
Dim objAttribute As IXMLDOMAttribute
Dim NoAtrs As Integer, AtrsName As String, AtrsVal As Variant
Dim i As Integer

'get File reference
If Not getDoc(xmlDoc, App.Path & "\XML_Dtd\db.xml") Then GoTo m_err:
MsgBox "No.of Attribute = " & xmlDoc.doctype.Attributes.length
For i = 0 To xmlDoc.doctype.Attributes.length - 1
  Set objAttribute = xmlDoc.doctype.Attributes.Item(i)
  MsgBox "Attribute Name=" & objAttribute.nodeName & vbCrLf & _
         "Attribute Value=" & objAttribute.nodeValue
Next i
m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End Sub

Private Sub Command2_Click()
'create the reader
Dim rdr As New SAXXMLReader40
'create the writer
Dim wrt As New MXXMLWriter40
Dim fileURL As String, xmlDoc As Msxml2.DOMDocument40

On Error GoTo errorHandler

wrt.byteOrderMark = True
wrt.omitXMLDeclaration = False
wrt.indent = True

'set the writer to the content handler
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt

'get File reference
If Not getDoc(xmlDoc, App.Path & "\books.xml") Then GoTo m_err:

'parse the XML
rdr.parse xmlDoc

'show the results in a message box, which can only display a maximum of around 1019 characters
MsgBox wrt.output

m_quit:
    Exit Sub
errorHandler:
    MsgBox Err.Description
    GoTo m_quit:
m_err:
    MsgBox "Error occurs!", vbCritical
    GoTo m_quit:

End Sub

Private Function getURL(tURL As String) As Boolean
Dim xmlDoc As New Msxml2.DOMDocument40
xmlDoc.async = False

If Not xmlDoc.Load(App.Path & "\books.xml") Then GoTo m_err:
tURL = xmlDoc.url
getURL = True

m_quit:
  Set xmlDoc = Nothing
  Exit Function
m_err:
  With xmlDoc.parseError
    Call displayError(.errorCode, .filePos, .line, .linePos, .reason, .srcText, .url)
  End With
  getURL = False
  GoTo m_quit:
End Function

Private Function getDoc(tDoc As Msxml2.DOMDocument40, URLParam As String) As Boolean
Dim xmlDoc As New Msxml2.DOMDocument40
xmlDoc.async = False

If Not xmlDoc.Load(URLParam) Then GoTo m_err:
Set tDoc = xmlDoc
getDoc = True

m_quit:
  Set xmlDoc = Nothing
  Exit Function
m_err:
  With xmlDoc.parseError
    Call displayError(.errorCode, .filePos, .line, .linePos, .reason, .srcText, .url)
  End With
  getDoc = False
  GoTo m_quit:
End Function

Private Sub test()

End Sub
Private Sub BuildTree()
    Dim i As Integer
    Dim TreeNode As Node                               ' node of treeview
    Dim xmlDoc As Msxml2.DOMDocument40
    Dim currentRootElement As Msxml2.IXMLDOMElement   'Mother
    Dim currentNodeElement As Msxml2.IXMLDOMNode      'Child
    Dim tempRootElement As Msxml2.IXMLDOMElement
    Dim tempNodeMap As Msxml2.IXMLDOMNamedNodeMap
    Dim objFunction As New ClsFunction
    
    'Reset the Treeview
    Me.TreeView1.Nodes.Clear
    
    Set xmlDoc = New Msxml2.DOMDocument40
    xmlDoc.async = False
    xmlDoc.Load (App.Path & "\books.xml")

    'The first element is always the Mother
    Set currentRootElement = xmlDoc.documentElement

    'add root to treeview
    Set tempNodeMap = currentRootElement.Attributes
         If tempNodeMap.length > 0 Then   'check if item has attribute
           Set tempRootElement = currentRootElement          'set current child as new mother
           ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
           If Not add2Tree(TreeView1, TreeNode, 0, 0, tempRootElement.getAttribute("id"), tempRootElement.nodeName) Then GoTo m_err:
         Else
           Set tempRootElement = currentRootElement          'set current child as new mother
            ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
           If Not add2Tree(TreeView1, TreeNode, 0, 0, , tempRootElement.nodeName) Then GoTo m_err:
         End If
        
    'check if current mother has child
    If currentRootElement.hasChildNodes Then
       
       TreeNode.Expanded = True   'Expand the root to display its child
       ' send to a method to perform loop
       Call objFunction.traverseNode(TreeView1, currentRootElement, TreeNode)
    End If
    Set objFunction = Nothing
         
m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End Sub

Private Sub Command3_Click()
Dim xmlDoc As Msxml2.DOMDocument40
Dim currentRootElement As Msxml2.IXMLDOMElement   'Mother
Dim currentNodeElement As Msxml2.IXMLDOMNode      'Child
Dim Temp As Msxml2.IXMLDOMAttribute
 
Set xmlDoc = New Msxml2.DOMDocument40
xmlDoc.async = False
xmlDoc.Load (App.Path & "\books.xml")
'Set root to the XML document's root element, COLLECTION:
Set currentRootElement = xmlDoc.documentElement
'Traverse through all child
For Each currentNodeElement In currentRootElement.childNodes
  MsgBox "Mother Element Name = " & currentRootElement.nodeName & vbCrLf & "Child Element Name = " & currentNodeElement.nodeName & vbCrLf & "Text = " & currentNodeElement.Text    'text and xml property r almost the same
Next

Set xmlDoc = Nothing
End Sub

Private Sub Command4_Click()
Call BuildTree
End Sub

Private Sub Command5_Click()
Dim i As Integer
Dim xmlDoc As Msxml2.DOMDocument40
Dim currentRootElement As Msxml2.IXMLDOMElement        'Mother
Dim CurrentNodeElementList As Msxml2.IXMLDOMNodeList   'Child collection
Dim objNodeList As IXMLDOMNodeList
 
Set xmlDoc = New Msxml2.DOMDocument40
xmlDoc.async = False
xmlDoc.Load (App.Path & "\books.xml")

Set objNodeList = xmlDoc.getElementsByTagName("book")
For i = 0 To (objNodeList.length - 1)
  MsgBox objNodeList.Item(i).Text
Next

m_quit:
  Set xmlDoc = Nothing
  Set currentRootElement = Nothing
  Exit Sub
m_err:
  GoTo m_quit:
End Sub



Private Sub Command6_Click(Index As Integer)
  Dim xmlDoc As Msxml2.DOMDocument40
  Dim currentRootElement As Msxml2.IXMLDOMElement   'Mother
  Dim currentNodeElement As Msxml2.IXMLDOMNode      'Child
  Dim Temp As Msxml2.IXMLDOMNamedNodeMap
  Dim temp2 As Msxml2.IXMLDOMElement
 
  Set xmlDoc = New Msxml2.DOMDocument40
  xmlDoc.async = False
  xmlDoc.Load (App.Path & "\books.xml")

If Index = 0 Then

'Set root to the XML document's root element, COLLECTION:
Set currentRootElement = xmlDoc.documentElement

  'Traverse through all child
  For Each currentNodeElement In currentRootElement.childNodes
    Set Temp = currentNodeElement.Attributes
  
    If Temp.length > 0 Then
       Set temp2 = currentNodeElement    'set current child as new mother
       MsgBox "Element Name : " & temp2.nodeName & vbCrLf & "Attribute Value :" & temp2.getAttribute("id")
       Set temp2 = Nothing
    End If
    Set Temp = Nothing
  Next
Else
 
  Dim objNodeList As IXMLDOMNodeList
  Dim nodenode As IXMLDOMNode
  Dim i As Integer, j As Integer
  
  Set objNodeList = xmlDoc.selectNodes("//book")
  For i = 0 To (objNodeList.length - 1)
    For j = 0 To objNodeList.Item(i).Attributes.length - 1
        Set nodenode = objNodeList.Item(i).Attributes.Item(j)
         MsgBox nodenode.nodeName & "=" & nodenode.nodeValue
    Next j
  Next i
End If

m_quit:
  Set xmlDoc = Nothing
  Set currentRootElement = Nothing
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Private Sub Command7_Click()
'create the reader
Dim rdr As New SAXXMLReader40
'create the writer
Dim wrt As New MXXMLWriter40
Dim fileURL As String, xmlDoc As Msxml2.DOMDocument40

On Error GoTo errorHandler

wrt.byteOrderMark = True
wrt.omitXMLDeclaration = False
wrt.indent = True

'set the writer to the content handler
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt

'get File reference
If Not getDoc(xmlDoc, App.Path & "\XML_Dtd\notes.xml") Then GoTo m_err:

'parse the XML
rdr.parse xmlDoc

'show the results in a message box, which can only display a maximum of around 1019 characters
MsgBox wrt.output

m_quit:
    Exit Sub
errorHandler:
    MsgBox Err.Description
    GoTo m_quit:
m_err:
    MsgBox "Error occurs!", vbCritical
    GoTo m_quit:

End Sub

Private Sub Command8_Click()
'create the reader
Dim rdr As New SAXXMLReader40
'create the writer
Dim wrt As New MXXMLWriter40
Dim fileURL As String, xmlDoc As Msxml2.DOMDocument40

On Error GoTo errorHandler

wrt.byteOrderMark = True
wrt.omitXMLDeclaration = False
wrt.indent = True

'set the writer to the content handler
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt

'get File reference
If Not getDoc(xmlDoc, App.Path & "\XML_Dtd\notes2.xml") Then GoTo m_err:

'parse the XML
rdr.parse xmlDoc

'show the results in a message box, which can only display a maximum of around 1019 characters
MsgBox wrt.output

m_quit:
    Exit Sub
errorHandler:
    MsgBox Err.Description
    GoTo m_quit:
m_err:
    MsgBox "Error occurs!", vbCritical
    GoTo m_quit:
End Sub

Private Sub Command9_Click()
'create the reader
Dim rdr As New SAXXMLReader40
'create the writer
Dim wrt As New MXXMLWriter40
Dim fileURL As String, xmlDoc As Msxml2.DOMDocument40

On Error GoTo errorHandler

wrt.byteOrderMark = True
wrt.omitXMLDeclaration = False
wrt.indent = True

'set the writer to the content handler
Set rdr.contentHandler = wrt
Set rdr.dtdHandler = wrt
Set rdr.errorHandler = wrt
rdr.putProperty "http://xml.org/sax/properties/lexical-handler", wrt
rdr.putProperty "http://xml.org/sax/properties/declaration-handler", wrt

'get File reference
If Not getDoc(xmlDoc, App.Path & "\XML_Dtd\notes3.xml") Then GoTo m_err:

'parse the XML
rdr.parse xmlDoc

'show the results in a message box, which can only display a maximum of around 1019 characters
MsgBox wrt.output

m_quit:
    Exit Sub
errorHandler:
    MsgBox Err.Description
    GoTo m_quit:
m_err:
    MsgBox "Error occurs!", vbCritical
    GoTo m_quit:
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
If Not Node.Key = "" Then MsgBox "ID = " & Node.Key
End Sub
