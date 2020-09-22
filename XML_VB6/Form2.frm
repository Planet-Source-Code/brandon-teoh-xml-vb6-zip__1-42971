VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   Caption         =   "Document Editor"
   ClientHeight    =   5445
   ClientLeft      =   720
   ClientTop       =   1470
   ClientWidth     =   10560
   LinkTopic       =   "Form2"
   ScaleHeight     =   5445
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "View External DTD File Contents"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   4920
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save File"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   4695
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xmlDoc As Msxml2.DOMDocument40
Dim DTDFilePath As String

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub execute(wrt As MXXMLWriter40, tdtdFilePath As String)
Text1.Text = wrt.output
DTDFilePath = tdtdFilePath

Me.Show vbModal

End Sub

Private Sub Command2_Click()
Dim fso As New FileSystemObject

On Error GoTo m_err2:
'test the validity of the xml document generated
If testDoc Then

    ' Set CancelError as True
    CommonDialog1.CancelError = True
    'set default path
    CommonDialog1.InitDir = App.Path
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    CommonDialog1.Filter = "All Files (*.*)|*.*|xml Files" & _
    "(*.xml)|*.xml|dtd Files (*.dtd)|*.dtd"
    ' Specify default filter
    CommonDialog1.FilterIndex = 2
    
    CommonDialog1.ShowSave 'display the dialog box with 'save' button
     
     
    With fso
        Set strm = .CreateTextFile(CStr(CommonDialog1.filename), True)
        strm.Write (Text1.Text)
    End With

    MsgBox "File Created Successfully!", vbInformation
Else
    GoTo m_err:
End If

m_quit:
   'delete tempFile
   Call Kill(App.Path & "\tempXML.tmp")
   Exit Sub
m_err:
   MsgBox "Please edit the document or DTD!", vbCritical
   GoTo m_quit:
m_err2:
   GoTo m_quit:
End Sub


Private Function testDoc() As Boolean
Dim fso As New FileSystemObject
testDoc = False

' save it to a temporarily file
 With fso
        Set strm = .CreateTextFile(App.Path & "/tempXML.tmp", True)
        strm.Write (Text1.Text)
 End With
    
Set strm = Nothing 'release its locking on tempXML.xml

' load it using parser
If Not parseXMLDoc(App.Path & "\tempXML.tmp") Then GoTo m_err:
testDoc = True

m_quit:
  Exit Function
m_err:
  testDoc = False
  GoTo m_quit:
End Function

Private Function parseXMLDoc(str As String) As Boolean
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
If Not getDoc(xmlDoc, str) Then GoTo m_err:

'parse the XML
rdr.parse xmlDoc
parseXMLDoc = True

m_quit:
    Exit Function
errorHandler:
    parseXMLDoc = False
    MsgBox Err.Description
    GoTo m_quit:
m_err:
    parseXMLDoc = False
    GoTo m_quit:
End Function

Private Function getDoc(tDoc As Msxml2.DOMDocument40, URLParam As String) As Boolean
Set xmlDoc = New Msxml2.DOMDocument40
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

Private Sub Command3_Click()
  If Not DTDFilePath = "" Then
     Call Form7.execute(DTDFilePath)
  Else
     MsgBox "Either there is no DTD file for this XML document or you didn't select DTD file!, go back and select", vbCritical
  End If
End Sub


