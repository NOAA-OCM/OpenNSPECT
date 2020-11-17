VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsXMLLUScenPollItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsXMLLUScenPollItem
' *************************************************************************************
' *  Description: XML Wrapper for use Pollutant Items with a management scenarios
' *
' *  Called By:
' *************************************************************************************
' *  Subs:
' *
' *
' *  Misc:
' *  Acknowledgements: Dave Grundgeiger       daveg@tarasoftware.com
' *                    Patrick Escarcega      patrick@vbguru.net
' *
' *************************************************************************************

Option Explicit

Private Const NODE_NAME = "LUScenPollutant"
Private Const ATTRIBUTE_PollID = "ID"
Private Const ELEMENT_Name = "PollName"
Private Const ELEMENT_Type1 = "Type1"
Private Const ELEMENT_Type2 = "Type2"
Private Const ELEMENT_Type3 = "Type3"
Private Const ELEMENT_Type4 = "Type4"

Public intID                    As Integer
Public strPollName              As String
Public intType1                 As Double
Public intType2                 As Double
Public intType3                 As Double
Public intType4                 As Double

Public Property Get NodeName() As String
'Retrieve the name of the element that this class wraps.
    NodeName = NODE_NAME
    
End Property

Public Property Get XML() As String
'Retrieve the XML string that this class represents. The XML returned is
'built from the values of this class's properties.
    
    XML = Me.CreateNode().XML

End Property

Public Property Let XML(ByVal strNewValue As String)
'Assign a new XML string to this class. The newly assigned XML is parsed,
'and the class's properties are set accordingly.
    
    Dim dom As New MSXML2.DOMDocument
    dom.loadXML strNewValue
    LoadNode dom.documentElement

End Property

Public Function CreateNode( _
    Optional Parent As MSXML2.IXMLDOMNode = Nothing _
) As MSXML2.IXMLDOMNode
'Return an XML DOM node that represents this class's properties. If a
'parent DOM node is passed in, then the returned node is also added as a
'child node of the parent.
        
    Dim node As MSXML2.IXMLDOMNode
    Dim dom As MSXML2.DOMDocument
    
    'If no parent was passed in, then create a DOM and document element.
    If Parent Is Nothing Then
        Set dom = New MSXML2.DOMDocument
        dom.loadXML "<" & NODE_NAME & "/>"
        Set node = dom.documentElement
    'Otherwise use passed-in parent.
    Else
        Set dom = Parent.ownerDocument
        Set node = dom.createElement(NODE_NAME)
        Parent.appendChild node
    End If
    
    NodeAppendAttribute dom, node, ATTRIBUTE_PollID, intID
    NodeAppendChildElement dom, node, ELEMENT_Name, strPollName
    NodeAppendChildElement dom, node, ELEMENT_Type1, intType1
    NodeAppendChildElement dom, node, ELEMENT_Type2, intType2
    NodeAppendChildElement dom, node, ELEMENT_Type3, intType3
    NodeAppendChildElement dom, node, ELEMENT_Type4, intType4
    
    'Return the created node
    Set CreateNode = node
    
End Function

Public Sub LoadNode(node As MSXML2.IXMLDOMNode)
'Set this class's properties based on the data found in the
'given node.

    'Ensure that a valid node was passed in.
    If node Is Nothing Then Exit Sub
    
    intID = GetNodeText(node, "@" & ATTRIBUTE_PollID)
    strPollName = GetNodeText(node, ELEMENT_Name)
    intType1 = CDbl(GetNodeText(node, ELEMENT_Type1))
    intType2 = CDbl(GetNodeText(node, ELEMENT_Type2))
    intType3 = CDbl(GetNodeText(node, ELEMENT_Type3))
    intType4 = CDbl(GetNodeText(node, ELEMENT_Type4))
        
End Sub

Private Function NodeGetText( _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal strNodeName As String _
) As String
'The NodeGetText function uses selectSingleNode on the passed-in node to
'find the child node having the given name. When found, the text of that
'node is returned. If the child node is not found in the node, then an
'empty string is returned.
    
    NodeGetText = "" 'default return value
    On Error Resume Next
    NodeGetText = node.selectSingleNode(strNodeName).Text
End Function

Private Sub NodeAppendAttribute( _
    ByVal dom As MSXML2.DOMDocument, _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal strName As String, _
    ByVal varValue As Variant _
)
'The NodeAppendAttribute subroutine creates an attribute having the given
'name and value, and adds it to the given node's Attributes collection.
    Dim attr As MSXML2.IXMLDOMAttribute
    
    'Create a new attribute and set its value.
    Set attr = dom.createAttribute(strName)
    attr.Value = varValue
    
    'Add the new attribute to the node's Attributes collection.
    Set attr = node.Attributes.setNamedItem(attr)
    attr.nodeValue = varValue
    
End Sub

Private Sub NodeAppendChildElement( _
    ByVal dom As MSXML2.DOMDocument, _
    ByVal node As MSXML2.IXMLDOMNode, _
    ByVal Name As String, _
    ByVal Value As Variant _
)
'The NodeAppendChildElement subroutine creates an element having the given
'name and value, and adds the element as a child of the given node.
    
    Dim element As MSXML2.IXMLDOMElement
    
    'Create a new child element and set its value.
    Set element = dom.createElement(Name)
    element.Text = CStr(Value)

    'Append the new child element to the node.
    node.appendChild element
    
End Sub

Public Function GetNodeText( _
    Parent As MSXML2.IXMLDOMNode, _
    ChildName As String _
) As String
'The GetNodeText function retrieves the value of the given child element
'within the given parent element. If the requested child element is not
'found, then an empty string is returned.

    Dim node As MSXML2.IXMLDOMNode
    
    On Error GoTo ErrorHandler
    
    GetNodeText = Parent.selectSingleNode(ChildName).Text
    Exit Function
    
ErrorHandler:
    
    GetNodeText = ""
    
End Function
