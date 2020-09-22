VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "MSXML4 API Demo"
   ClientHeight    =   3150
   ClientLeft      =   3390
   ClientTop       =   3915
   ClientWidth     =   2805
   LinkTopic       =   "Form3"
   ScaleHeight     =   3150
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit XML Document"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create XML Document"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "View XML Document"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
   Form1.Show vbModal
ElseIf Index = 1 Then
   Form4.Show vbModal
ElseIf Index = 2 Then
   Form6.Show vbModal
ElseIf Index = 3 Then
   Unload Me
End If

End Sub


