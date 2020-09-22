VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "DTD Editor"
   ClientHeight    =   4845
   ClientLeft      =   4635
   ClientTop       =   2430
   ClientWidth     =   6525
   LinkTopic       =   "Form7"
   ScaleHeight     =   4845
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject
Dim filePath As String

Public Sub execute(tfilePath As String)
filePath = tfilePath

With fso
  filePath = Replace(filePath, "file:///", "", 1, , vbTextCompare)
  filePath = LTrim(filePath)
  Text1.Text = ReadTextFileContents(filePath)
End With

Me.Show vbModal
End Sub

Function ReadTextFileContents(filename As String) As String
    Dim fnum As Integer, isOpen As Boolean
    On Error GoTo Error_Handler
    ' Get the next free file number.
    fnum = FreeFile()
    Open filename For Input As #fnum
    ' If execution flow got here, the file has been open without error.
    isOpen = True
    ' Read the entire contents in one single operation.
    ReadTextFileContents = Input(LOF(fnum), fnum)
    ' Intentionally flow into the error handler to close the file.
Error_Handler:
    ' Raise the error (if any), but first close the file.
    If isOpen Then Close #fnum
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

Private Sub Command1_Click(Index As Integer)
On Error GoTo m_err:
Dim strm As TextStream

If Index = 0 Then
  Unload Me
Else
  'save file
  Dim fso As New FileSystemObject
  
  With fso
       Set strm = .CreateTextFile(CStr(filePath), True)
       strm.Write (Text1.Text)
  End With

  MsgBox "File Updated!", vbInformation

m_quit:
   Exit Sub
m_err:
   GoTo m_quit:
End If

End Sub
