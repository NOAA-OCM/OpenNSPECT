Imports System.Xml

Public Class clsXMLCoeffTypeDef
    Inherits clsXMLBase
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsXMLCoeffTypeDef
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use Coefficient Type Definition
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


    'Following are the names of the NODES
    Private Const NODE_NAME As String = "TypeDefFile"
    Private Const NODE_TDLyrName As String = "TDLyrName" 'Layer Name
    Private Const NODE_TDLyrFileName As String = "TDLyrFileName" 'Layer FileName
    Private Const NODE_TDAttribute As String = "TDAttribute"
    Private Const NODE_TDType As String = "TDType" 'Alpha/Numeric 0 = Alpha, 1 = Numeric
    Private Const NODE_TDDef1 As String = "TDDef1"
    Private Const NODE_TDDef2 As String = "TDDef2"
    Private Const NODE_TDDef3 As String = "TDDef3"
    Private Const NODE_TDDef4 As String = "TDDef4"

    'Variables holding value of nodes above
    Public strTDLyrName As String
    Public strTDLyrFileName As String
    Public strTDAttribute As String
    Public intTDType As Short
    Public strTDDef1 As String
    Public strTDDef2 As String
    Public strTDDef3 As String
    Public strTDDef4 As String

    Public ReadOnly Property NodeName() As String
        Get
            'Retrieve the name of the element that this class wraps.

            NodeName = NODE_NAME

        End Get
    End Property



    Public Sub SaveFile(ByRef strXML As String)

        Dim dom As New XmlDocument
        dom.LoadXml(strXML)

        'TODO
        'frmPrj.grdCoeffs.set_TextMatrix(g_intCoeffRow, 6, strXML)

    End Sub


    Public Overrides Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        'Return an XML DOM node that represents this class's properties. If a
        'parent DOM node is passed in, then the returned node is also added as a
        'child node of the parent.

        Dim node As XmlNode
        Dim dom As XmlDocument

        'If no parent was passed in, then create a DOM and document element.
        If Parent Is Nothing Then
            dom = New XmlDocument
            dom.LoadXml("<" & NODE_NAME & "/>")
            node = dom.DocumentElement
            'Otherwise use passed-in parent.
        Else
            dom = Parent.OwnerDocument
            node = dom.CreateElement(NODE_NAME)
            Parent.AppendChild(node)
        End If

        '*********************************************************************
        NodeAppendChildElement(dom, node, NODE_TDLyrName, strTDLyrName)
        NodeAppendChildElement(dom, node, NODE_TDLyrFileName, strTDLyrFileName)
        NodeAppendChildElement(dom, node, NODE_TDAttribute, strTDAttribute)
        NodeAppendChildElement(dom, node, NODE_TDType, intTDType)
        NodeAppendChildElement(dom, node, NODE_TDDef1, strTDDef1)
        NodeAppendChildElement(dom, node, NODE_TDDef2, strTDDef2)
        NodeAppendChildElement(dom, node, NODE_TDDef3, strTDDef3)
        NodeAppendChildElement(dom, node, NODE_TDDef4, strTDDef4)

        'Return the created node

        CreateNode = node

    End Function


    Public Overrides Sub LoadNode(ByRef node As XmlNode)
        'Set this class's properties based on the data found in the
        'given node.
        'Ensure that a valid node was passed in.
        If node Is Nothing Then Exit Sub

        strTDLyrName = GetNodeText(node, NODE_TDLyrName)
        strTDLyrFileName = GetNodeText(node, NODE_TDLyrFileName)
        strTDAttribute = GetNodeText(node, NODE_TDAttribute)
        intTDType = CShort(GetNodeText(node, NODE_TDType))
        strTDDef1 = GetNodeText(node, NODE_TDDef1)
        strTDDef2 = GetNodeText(node, NODE_TDDef2)
        strTDDef3 = GetNodeText(node, NODE_TDDef3)
        strTDDef4 = GetNodeText(node, NODE_TDDef4)

    End Sub

End Class