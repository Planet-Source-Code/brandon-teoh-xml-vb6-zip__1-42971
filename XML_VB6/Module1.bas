Attribute VB_Name = "Module1"
Option Explicit
Public objXML As New ClsXML ' meant for ATTRIBUTE MAINTENANCE

Public Enum tSelect   ' enumeration
  myAttribute = 0
  myText = 1
End Enum


Public Function add2Tree(tTreeView As TreeView, Optional tTreeNode As Node, Optional tRelative As Integer, Optional tRelationship As Integer, Optional tKey As String, Optional tText As String) As Boolean

 If tRelative = 0 And tRelationship = 0 Then
    Set tTreeNode = tTreeView.Nodes.Add(, , tKey, tText) 'This line is meant for root element only
 Else
    Set tTreeNode = tTreeView.Nodes.Add(tRelative, tRelationship, tKey, tText)
 End If

add2Tree = True

m_quit:
  Exit Function
m_err:
  add2Tree = False
  GoTo m_quit:
End Function

Public Function add2TreeWithImage(tTreeView As TreeView, Optional tTreeNode As Node, Optional tRelative As Integer, Optional tRelationship As Integer, Optional tKey As String, Optional tText As String, Optional tImage As Variant, Optional tSelectedImage As Variant) As Boolean
 If tRelative = 0 And tRelationship = 0 Then
    Set tTreeNode = tTreeView.Nodes.Add(, , tKey, tText, tImage, tSelectedImage)   'This line is meant for root element only
    If tImage = 1 Then
       tTreeNode.Tag = "NODE"  'set node type
    Else
       tTreeNode.Tag = "TEXT"  'set node type
    End If
 Else
    Set tTreeNode = tTreeView.Nodes.Add(tRelative, tRelationship, tKey, tText, tImage, tSelectedImage)
    If tImage = 1 Then
       tTreeNode.Tag = "NODE"  'set node type
     Else
       tTreeNode.Tag = "TEXT"  'set node type
    End If
 End If

add2TreeWithImage = True

m_quit:
  Exit Function
m_err:
  add2TreeWithImage = False
  GoTo m_quit:
End Function

Public Sub exploreAttribute(tCurrentNode As Msxml2.IXMLDOMNode, Optional tTreeNode As Node)
On Error GoTo m_err:

Dim ATT_Name As String, ATT_Val As String, myKey As String
Dim currentRootElement As Msxml2.IXMLDOMElement
Dim atrsNode As IXMLDOMNode
Dim objATTNameValue As ClsATTNameValue
Dim objAttribute As New ClsAttribute
Dim i As Integer

Set objAttribute.myNode = tTreeNode

If tCurrentNode.nodeType = 1 Then
   Set currentRootElement = tCurrentNode
   For i = 0 To (currentRootElement.Attributes.length - 1)
      Set atrsNode = currentRootElement.Attributes.Item(i)
      If atrsNode.specified Then     'make sure current attribute is explicitly specified
         ATT_Name = atrsNode.nodeName
         ATT_Val = atrsNode.nodeValue
         'myKey = tTreeNode.index & "_" & ATT_Name & "_" & ATT_Val
         myKey = ATT_Name
         Set objATTNameValue = New ClsATTNameValue
         objATTNameValue.myKey = myKey
         objATTNameValue.ATT_Name = ATT_Name
         objATTNameValue.ATT_Value = ATT_Val
         'add objATTNamevalue to objAttribute
         Call objAttribute.AddItem(objATTNameValue, myKey)
         Set objATTNameValue = Nothing
      End If
    Next i
Else
   '-------------DO NOTHING HERE ----------------------'
End If

'add objAttribute to objxml
Call objXML.AddItem(myAttribute, objAttribute)
Set objAttribute = Nothing

m_quit:
  Exit Sub
m_err:
  MsgBox Err.Description
  GoTo m_quit:
End Sub

Public Sub displayError(errorCode As Variant, filePos As Variant, line As Variant, linePos As Variant, reason As Variant, srcText As Variant, url As Variant)

MsgBox "Error Code : " & errorCode & vbCrLf & _
       "File Position : " & filePos & vbCrLf & _
       "Line : " & line & vbCrLf & _
       "Line Position : " & linePos & vbCrLf & _
       "Reason : " & reason & vbCrLf & _
       "Src Text : " & srcText & vbCrLf & _
       "URL : " & url, vbCritical
End Sub

