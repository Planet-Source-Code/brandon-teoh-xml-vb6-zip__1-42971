VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This method is used to push treeview element from XML Doc to treeview
Public Sub traverseNode(tTreeView As TreeView, currentRootElement As Msxml2.IXMLDOMElement, tTreeNode As Node)
On Error Resume Next  'exception handling
Dim NewTreeNode As Node
Dim currentNodeElement As Msxml2.IXMLDOMNode
Dim tempRootElement As Msxml2.IXMLDOMElement
Dim tempNodeMap As Msxml2.IXMLDOMNamedNodeMap
Dim objFunction As ClsFunction
    
For Each currentNodeElement In currentRootElement.childNodes
    If currentNodeElement.nodeType = 1 Then   '1 = NODE_ELEMENT, 2 = NODE_ATTRIBUTE, 3 = NODE_TEXT
       Set tempNodeMap = currentNodeElement.Attributes
       If tempNodeMap.length > 0 Then
           Set tempRootElement = currentNodeElement          'set current child as new mother
           ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
           If Not add2Tree(tTreeView, NewTreeNode, tTreeNode.Index, 4, tempRootElement.getAttribute("id"), tempRootElement.nodeName) Then GoTo m_err:
        Else
           Set tempRootElement = currentNodeElement          'set current child as new mother
            ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
           If Not add2Tree(tTreeView, NewTreeNode, tTreeNode.Index, 4, , tempRootElement.nodeName) Then GoTo m_err:
        End If
        If currentNodeElement.hasChildNodes Then
            ' --------------------------RECURSIVE CALL-------------------------------------------------------------------------------------------
            Set objFunction = New ClsFunction   ' create new object so that the next process is create at a new stack, preventing stack overflow
            Call objFunction.traverseNode(tTreeView, tempRootElement, NewTreeNode)
            ' --------------------------END OF RECURSION -----------------------------------------------------------------------------------------
        End If
    Else
            ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
           If Not add2Tree(tTreeView, NewTreeNode, tTreeNode.Index, 4, , currentNodeElement.Text) Then GoTo m_err:
    End If
    Set tempNodeMap = Nothing
    Set tempRootElement = Nothing
    Set objFunction = Nothing
Next

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub
'This method is being used to push treeview element from treeview to xml Doc
Public Sub traverseNode2(currentNode As Node)
Dim i As Integer
Dim objFunction As ClsFunction
Dim totalNode As Integer
Dim tempNode As Node

' get the number of children
totalNode = currentNode.children
Set tempNode = currentNode.Child

For i = 1 To totalNode
   Debug.Print tempNode.Text
      If tempNode.children <> 0 Then
    '---------------RECURSIVE START HERE ------------------'
        Set objFunction = New ClsFunction
        Call objFunction.traverseNode2(tempNode)
    '---------------RECURSIVE END HERE --------------------'
      End If
   Set objFunction = Nothing
   Set tempNode = tempNode.Next
Next i

End Sub

'This method is used by form4
Public Sub traverseNode2XML(currentNode As Node, ByRef tCntH As IVBSAXContentHandler)
Dim i As Integer
Dim objFunction As ClsFunction
Dim totalNode As Integer
Dim tempNode As Node
Dim atrs As New SAXAttributes40
Dim Cur_Str As String

' get the number of children
totalNode = currentNode.children
Set tempNode = currentNode.Child

For i = 1 To totalNode
   ' get attributes of current node
   Call getAttributes(atrs, tempNode.Index)  'we always start with the first child, thus has to add 1 to the current loop counter
   
   tCntH.startElement "", "", tempNode.Text, atrs
   
   'Get the element of current node
   If getText(Cur_Str, tempNode.Index) Then tCntH.characters Cur_Str  'add to contenthandlers
           
      If tempNode.children <> 0 Then
    '---------------RECURSIVE START HERE ------------------'
        Set objFunction = New ClsFunction
        Call objFunction.traverseNode2XML(tempNode, tCntH)
    '---------------RECURSIVE END HERE --------------------'
      End If
   Set objFunction = Nothing
   Set atrs = Nothing
   tCntH.endElement "", "", tempNode.Text
   Set tempNode = tempNode.Next
Next i

End Sub

'This method is slightly different from traverseNode2XML, being used by form6
Public Sub traverseNode2XML2(currentNode As Node, ByRef tCntH As IVBSAXContentHandler)
Dim i As Integer
Dim objFunction As ClsFunction
Dim totalNode As Integer
Dim tempNode As Node
Dim atrs As New SAXAttributes40
Dim Cur_Str As String

' get the number of children
totalNode = currentNode.children
Set tempNode = currentNode.Child

For i = 1 To totalNode
   If tempNode.Tag = "TEXT" Then   ' check if current treenode is a text or a node
      Call tCntH.characters(tempNode.Text)    'add text to contenthandlers
      
      If tempNode.children <> 0 Then
      '---------------RECURSIVE START HERE ------------------'
        Set objFunction = New ClsFunction
        Call objFunction.traverseNode2XML2(tempNode, tCntH)
      '---------------RECURSIVE END HERE --------------------'
        Set objFunction = Nothing
      End If
   Else
      ' get attributes of current node
      Call getAttributes(atrs, tempNode.Index)  'we always start with the first child, thus has to add 1 to the current loop counter
   
      tCntH.startElement "", "", tempNode.Text, atrs
           
      If tempNode.children <> 0 Then
      '---------------RECURSIVE START HERE ------------------'
        Set objFunction = New ClsFunction
        Call objFunction.traverseNode2XML2(tempNode, tCntH)
      '---------------RECURSIVE END HERE --------------------'
        Set objFunction = Nothing
      End If
      tCntH.endElement "", "", tempNode.Text
      Set atrs = Nothing
   End If
   Set tempNode = tempNode.Next
Next i

End Sub
Private Sub getAttributes(ByRef tAtrs As SAXAttributes40, Index As Integer)
     Dim objAttribute As ClsAttribute
     Dim objATTNameValue As ClsATTNameValue
     Dim j As Integer
     
           Set objAttribute = objXML.getItem_Index(myAttribute, Index)
           For j = 1 To objAttribute.getItemCount
               Set objATTNameValue = objAttribute.getItem_Index(j)
               tAtrs.addAttribute "", "", objATTNameValue.ATT_Name, "", objATTNameValue.ATT_Value
           Next j
     
     Set objAttribute = Nothing
     Set objATTNameValue = Nothing
End Sub

Private Function getText(ByRef str As String, Index As Integer) As Boolean
   Dim ObjText As ClsText
   getText = False
      Set ObjText = objXML.getItem_Index(myText, Index)
      str = ObjText.myData
      getText = True
      Exit Function

   Set ObjText = Nothing
End Function

Public Sub traverseNodeWithImage(tTreeView As TreeView, currentRootElement As Msxml2.IXMLDOMElement, tTreeNode As Node)
On Error Resume Next  'exception handling
Dim NewTreeNode As Node
Dim currentNodeElement As Msxml2.IXMLDOMNode
Dim tempRootElement As Msxml2.IXMLDOMElement
Dim tempNodeMap As Msxml2.IXMLDOMNamedNodeMap
Dim objFunction As ClsFunction

For Each currentNodeElement In currentRootElement.childNodes
    If currentNodeElement.nodeType = 1 Then   '1 = NODE_ELEMENT, 2 = NODE_ATTRIBUTE, 3 = NODE_TEXT
        
        'add current node to treeview
        If Not add2TreeWithImage(tTreeView, NewTreeNode, tTreeNode.Index, 4, , currentNodeElement.nodeName, 1, 2) Then GoTo m_err:  ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
        Call exploreAttribute(currentNodeElement, NewTreeNode)
        
        'explore child
        If currentNodeElement.hasChildNodes Then
           '----------------------------Perform Recursive ----------------------------------'
            Set objFunction = New ClsFunction
            Call objFunction.traverseNodeWithImage(tTreeView, currentNodeElement, NewTreeNode)
           '---------------------------------------------------------------------------------'
        End If
    ElseIf currentNodeElement.nodeType = 2 Then '1 = NODE_ELEMENT, 2 = NODE_ATTRIBUTE, 3 = NODE_TEXT
           '---------------------This segment is never performed ---------------------------'
    ElseIf currentNodeElement.nodeType = 3 Then '1 = NODE_ELEMENT, 2 = NODE_ATTRIBUTE, 3 = NODE_TEXT
           If Not add2TreeWithImage(tTreeView, NewTreeNode, tTreeNode.Index, 4, , currentNodeElement.Text, 3) Then GoTo m_err: ' 1- tvwLast, 2 - tvwNext, 3-tvwPrevious, 4-tvwChild
           Call exploreAttribute(currentNodeElement, NewTreeNode)
    
           'explore child
           If currentNodeElement.hasChildNodes Then
           '----------------------------Perform Recursive ----------------------------------'
            Set objFunction = New ClsFunction
            Call objFunction.traverseNodeWithImage(tTreeView, currentNodeElement, NewTreeNode)
           '---------------------------------------------------------------------------------'
        End If
        
    End If
    Set tempNodeMap = Nothing
    Set tempRootElement = Nothing
    Set objFunction = Nothing
Next

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Public Sub scanChild(ByRef tCol As Collection, ByVal currentNode As Node)

Dim i As Integer
Dim objFunction As ClsFunction
Dim totalNode As Integer
Dim tempNode As Node
Dim objGeneral As New ClsGeneral

' get the number of children
totalNode = currentNode.children
Set tempNode = currentNode.Child

 For i = 1 To totalNode
      If tempNode.children <> 0 Then
      '---------------RECURSIVE START HERE ------------------'
         Set objFunction = New ClsFunction  ' create new object so that the next process is create at a new stack, preventing stack overflow
         Call objFunction.scanChild(tCol, tempNode)
      '---------------RECURSIVE END HERE --------------------'
         objGeneral.myValue = tempNode.Index
         Call tCol.Add(objGeneral)  ' add object to collection
      Else
         objGeneral.myValue = tempNode.Index
         Call tCol.Add(objGeneral)  ' add object to collection
      End If
      Set objFunction = Nothing
      Set objGeneral = Nothing
      Set tempNode = tempNode.Next
 Next i

m_quit:
  Exit Sub
m_err:
  GoTo m_quit:
End Sub

Public Sub calTimeDiff(startValue As Double, ByRef tYears As Integer, ByRef tMonths As Integer, ByRef tDays As Integer, ByRef tHours As Integer, ByRef tMinutes As Integer, ByRef tSeconds As Integer, tPreviousLevel As String)
Dim objFunction As ClsFunction
Dim tempVal As Double, fixVal As Integer, multiplyConst As Integer, currentLevel As String

If tPreviousLevel = "MINUTE" Then
   multiplyConst = 60
   tSeconds = startValue * multiplyConst
Else
  If tPreviousLevel = "YEAR" Then
     multiplyConst = 12
     currentLevel = "MONTH"
     tempVal = startValue * multiplyConst
     tMonths = Fix(tempVal)           ' get the non-decimal portion
     fixVal = tMonths                 'set reference
  ElseIf tPreviousLevel = "MONTH" Then
     multiplyConst = 30   'average 30 days per month
     currentLevel = "DAY"             'set reference
     tempVal = startValue * multiplyConst
     tDays = Fix(tempVal)           ' get the non-decimal portion
     fixVal = tDays                 'set reference
  ElseIf tPreviousLevel = "DAY" Then
     multiplyConst = 24
     currentLevel = "HOUR"
     tempVal = startValue * multiplyConst
     tHours = Fix(tempVal)           ' get the non-decimal portion
     fixVal = tHours                 'set reference
  ElseIf tPreviousLevel = "HOUR" Then
     multiplyConst = 60
     currentLevel = "MINUTE"
     tempVal = startValue * multiplyConst
     tMinutes = Fix(tempVal)           ' get the non-decimal portion
     fixVal = tMinutes                 'set reference
  End If
  startValue = Abs(tempVal - fixVal)   'extract the absolute value of the decimal portion
  If startValue > 0 Then
     Set objFunction = New ClsFunction
     Call objFunction.calTimeDiff(startValue, tYears, tMonths, tDays, tHours, tMinutes, tSeconds, currentLevel)
  End If
  Set objFunction = Nothing
End If

End Sub

