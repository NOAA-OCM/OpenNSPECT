Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsXMLPollutantItem_NET.clsXMLPollutantItem")> Public Class clsXMLPollutantItem
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsXMLPollutantItem
	' *************************************************************************************
	' *  Description: XML Wrapper for use a single pollutant in the frmPrj Pollutants Tab
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
	
	
	Private Const NODE_NAME As String = "Pollutant"
	Private Const ATTRIBUTE_ID As String = "ID"
	Private Const ELEMENT_Apply As String = "Apply"
	Private Const ELEMENT_PollName As String = "PollName"
	Private Const ELEMENT_CoeffSet As String = "CoeffSet"
	Private Const ELEMENT_Coeff As String = "Coeff"
	Private Const ELEMENT_Threshold As String = "Threshold"
	Private Const ELEMENT_TypeDefXMLFile As String = "TypeDefXMLFile"
	
	Public intID As Short
	Public intApply As Short
	Public strPollName As String
	Public strCoeffSet As String
	Public strCoeff As String
	Public intThreshold As Short
	Public strTypeDefXMLFile As String
	
	Public ReadOnly Property NodeName() As String
		Get
			'Retrieve the name of the element that this class wraps.
			
			NodeName = NODE_NAME
			
		End Get
	End Property
	
	
	Public Property XML() As String
		Get
			'Retrieve the XML string that this class represents. The XML returned is
			'built from the values of this class's properties.
			
			XML = Me.CreateNode().XML
			
		End Get
		Set(ByVal Value As String)
			'Assign a new XML string to this class. The newly assigned XML is parsed,
			'and the class's properties are set accordingly.
			
			Dim dom As New MSXML2.DOMDocument
			dom.loadXML(Value)
			'UPGRADE_WARNING: Couldn't resolve default property of object dom.documentElement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LoadNode((dom.documentElement))
			
		End Set
	End Property
	
	Public Function CreateNode(Optional ByRef Parent As MSXML2.IXMLDOMNode = Nothing) As MSXML2.IXMLDOMNode
		'Return an XML DOM node that represents this class's properties. If a
		'parent DOM node is passed in, then the returned node is also added as a
		'child node of the parent.
		
		Dim node As MSXML2.IXMLDOMNode
		Dim dom As MSXML2.DOMDocument
		
		'If no parent was passed in, then create a DOM and document element.
		If Parent Is Nothing Then
			dom = New MSXML2.DOMDocument
			dom.loadXML("<" & NODE_NAME & "/>")
			node = dom.documentElement
			'Otherwise use passed-in parent.
		Else
			dom = Parent.ownerDocument
			node = dom.createElement(NODE_NAME)
			Parent.appendChild(node)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		node.appendChild(dom.createTextNode(vbNewLine & vbTab))
		NodeAppendAttribute(dom, node, ATTRIBUTE_ID, intID)
		NodeAppendChildElement(dom, node, ELEMENT_Apply, intApply)
		NodeAppendChildElement(dom, node, ELEMENT_PollName, strPollName)
		NodeAppendChildElement(dom, node, ELEMENT_CoeffSet, strCoeffSet)
		NodeAppendChildElement(dom, node, ELEMENT_Coeff, strCoeff)
		NodeAppendChildElement(dom, node, ELEMENT_Threshold, intThreshold)
		NodeAppendChildElement(dom, node, ELEMENT_TypeDefXMLFile, strTypeDefXMLFile)
		'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		node.appendChild(dom.createTextNode(vbNewLine & vbTab))
		
		'Return the created node
		CreateNode = node
		
	End Function
	
	Public Sub LoadNode(ByRef node As MSXML2.IXMLDOMNode)
		'Set this class's properties based on the data found in the
		'given node.
		
		'Ensure that a valid node was passed in.
		If node Is Nothing Then Exit Sub
		
		intID = CShort(GetNodeText(node, "@" & ATTRIBUTE_ID))
		intApply = CShort(GetNodeText(node, ELEMENT_Apply))
		strPollName = GetNodeText(node, ELEMENT_PollName)
		strCoeffSet = GetNodeText(node, ELEMENT_CoeffSet)
		strCoeff = GetNodeText(node, ELEMENT_Coeff)
		intThreshold = CShort(GetNodeText(node, ELEMENT_Threshold))
		strTypeDefXMLFile = GetNodeText(node, ELEMENT_TypeDefXMLFile)
		
	End Sub
	
	Private Function NodeGetText(ByVal node As MSXML2.IXMLDOMNode, ByVal strNodeName As String) As String
		'The NodeGetText function uses selectSingleNode on the passed-in node to
		'find the child node having the given name. When found, the text of that
		'node is returned. If the child node is not found in the node, then an
		'empty string is returned.
		
		NodeGetText = "" 'default return value
		On Error Resume Next
		NodeGetText = node.selectSingleNode(strNodeName).Text
		
	End Function
	
	Private Sub NodeAppendAttribute(ByVal dom As MSXML2.DOMDocument, ByVal node As MSXML2.IXMLDOMNode, ByVal strName As String, ByVal varValue As Object)
		'The NodeAppendAttribute subroutine creates an attribute having the given
		'name and value, and adds it to the given node's Attributes collection.
		
		Dim attr As MSXML2.IXMLDOMAttribute
		
		'Create a new attribute and set its value.
		attr = dom.createAttribute(strName)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		attr.Value = varValue
		
		'Add the new attribute to the node's Attributes collection.
		'UPGRADE_WARNING: Couldn't resolve default property of object attr. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		attr = node.Attributes.setNamedItem(attr)
		'UPGRADE_WARNING: Couldn't resolve default property of object varValue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		attr.nodeValue = varValue
		
	End Sub
	
	Private Sub NodeAppendChildElement(ByVal dom As MSXML2.DOMDocument, ByVal node As MSXML2.IXMLDOMNode, ByVal Name As String, ByVal Value As Object)
		'The NodeAppendChildElement subroutine creates an element having the given
		'name and value, and adds the element as a child of the given node.
		
		Dim element As MSXML2.IXMLDOMElement
		
		'Create a new child element and set its value.
		element = dom.createElement(Name)
		'UPGRADE_WARNING: Couldn't resolve default property of object Value. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		element.Text = CStr(Value)
		
		'Append the new child element to the node.
		'UPGRADE_WARNING: Couldn't resolve default property of object element. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		node.appendChild(element)
		'Format nicely
		'UPGRADE_WARNING: Couldn't resolve default property of object dom.createTextNode(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		node.appendChild(dom.createTextNode(vbNewLine & vbTab))
		
	End Sub
	
	Public Function GetNodeText(ByRef Parent As MSXML2.IXMLDOMNode, ByRef ChildName As String) As String
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
End Class