VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mother & Child"
   ClientHeight    =   945
   ClientLeft      =   3255
   ClientTop       =   1920
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   2385
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'nodeType : 1 = element, 3 = xml-values

 Dim xmlDoc As MSXML2.DOMDocument40
 Dim currentRootElement As MSXML2.IXMLDOMElement, tempRootElement As MSXML2.IXMLDOMElement, tempRootElement2 As MSXML2.IXMLDOMElement    'Mother
 Dim currentNodeElement As MSXML2.IXMLDOMNode, tempNodeElement As MSXML2.IXMLDOMNode, tempNodeElement2 As MSXML2.IXMLDOMNode              'Child
    
 Set xmlDoc = New MSXML2.DOMDocument40
 xmlDoc.async = False
 xmlDoc.Load (App.Path & "\books2.xml")

 'The first element is always the Mother
 Set currentRootElement = xmlDoc.documentElement
    
 If currentRootElement.hasChildNodes Then
    If currentRootElement.nodeType = 3 Then
       MsgBox getNodeType(currentRootElement.nodeType) & " '" & currentRootElement.Text & "' is childless", vbInformation
    Else
       MsgBox getNodeType(currentRootElement.nodeType) & " '" & currentRootElement.nodeName & "' has child(s)", vbInformation
    End If
    For Each currentNodeElement In currentRootElement.childNodes
        If currentNodeElement.hasChildNodes Then
           If currentNodeElement.nodeType = 3 Then
              MsgBox getNodeType(currentNodeElement.nodeType) & " '" & currentNodeElement.Text & "' is childless", vbInformation
           Else
              MsgBox getNodeType(currentNodeElement.nodeType) & " '" & currentNodeElement.nodeName & "' has child(s)", vbInformation
           End If
           Set tempRootElement = currentNodeElement
               For Each tempNodeElement In tempRootElement.childNodes
                   If tempNodeElement.hasChildNodes Then
                      If tempNodeElement.nodeType = 3 Then
                         MsgBox getNodeType(tempNodeElement.nodeType) & " '" & tempNodeElement.Text & "' is childless", vbInformation
                      Else
                         MsgBox getNodeType(tempNodeElement.nodeType) & " '" & tempNodeElement.nodeName & "' has child(s)", vbInformation
                      End If
                      Set tempRootElement2 = tempNodeElement
                      For Each tempNodeElement2 In tempRootElement2.childNodes
                          If tempNodeElement2.hasChildNodes Then
                             MsgBox "We shouldn't be here!", vbCritical
                          Else
                             If tempNodeElement2.nodeType = 3 Then
                                MsgBox getNodeType(tempNodeElement2.nodeType) & " '" & tempNodeElement2.Text & "' is childless", vbInformation
                             Else
                                MsgBox getNodeType(tempNodeElement2.nodeType) & " '" & tempNodeElement2.nodeName & "' is childless", vbInformation
                             End If
                          End If
                      Next tempNodeElement2
                   Else
                      MsgBox getNodeType(tempNodeElement.nodeType) & " '" & tempNodeElement.nodeName & "' is childless", vbInformation
                   End If
               Next tempNodeElement
        Else
           MsgBox getNodeType(currentNodeElement.nodeType) & " '" & currentNodeElement.nodeType & "' is childless", vbInformation
        End If
    Next currentNodeElement
 Else
    MsgBox getNodeType(currentRootElement.nodeType) & " '" & currentRootElement.nodeType & " is childless", vbInformation
 End If

End Sub

Private Function getNodeType(i As Integer) As String

If i = 1 Then
   getNodeType = "NODE_ELEMENT"
ElseIf i = 2 Then
   getNodeType = "NODE_ATTRIBUTE"
ElseIf i = 3 Then
   getNodeType = "NODE_TEXT"
ElseIf i = 4 Then
   getNodeType = "NODE_CDATA_SECTION"
ElseIf i = 5 Then
   getNodeType = "NODE_ENTITY_REFERENCE"
ElseIf i = 6 Then
   getNodeType = "NODE_ENTITY"
ElseIf i = 7 Then
   getNodeType = "NODE_PROCESSING_INSTRUCTION"
ElseIf i = 8 Then
   getNodeType = "NODE_COMMENT"
ElseIf i = 9 Then
   getNodeType = "NODE_DOCUMENT"
ElseIf i = 10 Then
   getNodeType = "NODE_DOCUMENT_TYPE"
ElseIf i = 11 Then
   getNodeType = "NODE_DOCUMENT_FRAGMENT"
ElseIf i = 12 Then
   getNodeType = "NODE_NOTATION"
End If

End Function
