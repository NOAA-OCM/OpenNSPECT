'********************************************************************************************************
'File Name: clsWrapperManScen.vb
'Description: A class to handle the management scenario xml
'********************************************************************************************************
'The contents of this file are subject to the Mozilla Public License Version 1.1 (the "License"); 
'you may not use this file except in compliance with the License. You may obtain a copy of the License at 
'http://www.mozilla.org/MPL/ 
'Software distributed under the License is distributed on an "AS IS" basis, WITHOUT WARRANTY OF 
'ANY KIND, either express or implied. See the License for the specificlanguage governing rights and 
'limitations under the License. 
'
'Note: This code was converted from the vb6 NSPECT ArcGIS extension and so bears many of the old comments
'in the files where it was possible to leave them.
'
'Contributor(s): (Open source contributors should list themselves and their modifications here). 
'Oct 20, 2010:  Allen Anselmo allen.anselmo@gmail.com - 
'               Added licensing and comments to code
Imports System.Collections.Generic
Imports System.Xml

Public Class clsXMLLUScen
    Inherits clsXMLBase
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsWrapperMain
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use Management scenarios
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

    Private Const NODE_NAME As String = "LUScenFile"
    Private Const NODE_ManScenName As String = "LUScenName"
    Private Const NODE_ManScenLyrName As String = "LUScenLyrName"
    Private Const NODE_ManScenFileName As String = "LUScenFileName"
    Private Const NODE_ManScenSelectedPoly As String = "LUScenSelectedPoly"
    Private Const NODE_ManScenSelectedPolyList As String = "LUScenSelectedPolyList"
    Private Const NODE_SCSCurveA As String = "SCSCurveA"
    Private Const NODE_SCSCurveB As String = "SCSCurveB"
    Private Const NODE_SCSCurveC As String = "SCSCurveC"
    Private Const NODE_SCSCurveD As String = "SCSCurveD"
    Private Const NODE_CoverFactor As String = "CoverFactor"
    Private Const NODE_WaterWetlands As String = "WaterWetlands"

    'Variables holding value of nodes above
    Public strLUScenName As String
    Public strLUScenLyrName As String
    Public strLUScenFileName As String
    Public intLUScenSelectedPoly As Short
    Public intLUScenSelectedPolyList As New List(Of Integer)
    Public intSCSCurveA As Double
    Public intSCSCurveB As Double
    Public intSCSCurveC As Double
    Public intSCSCurveD As Double
    Public lngCoverFactor As Double
    Public intWaterWetlands As Short

    Public clsPollutant As clsXMLLUScenPollItem
    Public clsPollItems As clsXMLLUScenPollItems

    Public ReadOnly Property NodeName() As String
        Get
            'Retrieve the name of the element that this class wraps.
            NodeName = NODE_NAME

        End Get
    End Property

    Public Overrides Function CreateNode (Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        Try
            'Return an XML DOM node that represents this class's properties. If a
            'parent DOM node is passed in, then the returned node is also added as a
            'child node of the parent.

            Dim node As XmlNode
            Dim dom As XmlDocument

            'If no parent was passed in, then create a DOM and document element.
            If Parent Is Nothing Then
                dom = New XmlDocument
                dom.LoadXml ("<" & NODE_NAME & "/>")
                node = dom.DocumentElement
                'Otherwise use passed-in parent.
            Else
                dom = Parent.OwnerDocument
                node = dom.CreateElement (NODE_NAME)
                Parent.AppendChild (node)
            End If

            '*********************************************************************
            NodeAppendChildElement (dom, node, NODE_ManScenName, strLUScenName)
            NodeAppendChildElement (dom, node, NODE_ManScenLyrName, strLUScenLyrName)
            NodeAppendChildElement (dom, node, NODE_ManScenFileName, strLUScenFileName)
            NodeAppendChildElement (dom, node, NODE_ManScenSelectedPoly, intLUScenSelectedPoly)
            Dim strlist As String = ""
            If intLUScenSelectedPolyList.Count > 0 Then strlist = intLUScenSelectedPolyList (0).ToString
            For i As Integer = 1 To intLUScenSelectedPolyList.Count - 1
                strlist = strlist + "," + intLUScenSelectedPolyList (i).ToString
            Next
            NodeAppendChildElement (dom, node, NODE_ManScenSelectedPolyList, strlist)
            NodeAppendChildElement (dom, node, NODE_SCSCurveA, intSCSCurveA)
            NodeAppendChildElement (dom, node, NODE_SCSCurveB, intSCSCurveB)
            NodeAppendChildElement (dom, node, NODE_SCSCurveC, intSCSCurveC)
            NodeAppendChildElement (dom, node, NODE_SCSCurveD, intSCSCurveD)
            NodeAppendChildElement (dom, node, NODE_CoverFactor, lngCoverFactor)
            NodeAppendChildElement (dom, node, NODE_WaterWetlands, intWaterWetlands)

            clsPollItems.CreateNode (node)

            CreateNode = node

        Catch ex As Exception
            HandleError (ex)
            CreateNode = Nothing
        End Try
    End Function

    Public Overrides Sub LoadNode (ByRef node As XmlNode)
        Try
            'Set this class's properties based on the data found in the
            'given node.

            'Ensure that a valid node was passed in.
            If node Is Nothing Then Exit Sub

            strLUScenName = GetNodeText (node, NODE_ManScenName)
            strLUScenLyrName = GetNodeText (node, NODE_ManScenLyrName)
            strLUScenFileName = GetNodeText (node, NODE_ManScenFileName)
            intLUScenSelectedPoly = CShort (GetNodeText (node, NODE_ManScenSelectedPoly))
            Dim tmpstr As String() = GetNodeText (node, NODE_ManScenSelectedPolyList).Split (",")
            intLUScenSelectedPolyList.Clear()
            For i As Integer = 0 To tmpstr.Length - 1
                If tmpstr (i) <> "" Then
                    intLUScenSelectedPolyList.Add (CShort (tmpstr (i)))
                End If
            Next
            intSCSCurveA = CDbl (GetNodeText (node, NODE_SCSCurveA))
            intSCSCurveB = CDbl (GetNodeText (node, NODE_SCSCurveB))
            intSCSCurveC = CDbl (GetNodeText (node, NODE_SCSCurveC))
            intSCSCurveD = CDbl (GetNodeText (node, NODE_SCSCurveD))
            lngCoverFactor = CDbl (GetNodeText (node, NODE_CoverFactor))
            intWaterWetlands = CShort (GetNodeText (node, NODE_WaterWetlands))

            clsPollItems.LoadNode (node.SelectSingleNode (clsPollItems.NodeName))

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub

    Public Sub New()
        Try
            clsPollutant = New clsXMLLUScenPollItem
            clsPollItems = New clsXMLLUScenPollItems

        Catch ex As Exception
            HandleError (ex)
        End Try
    End Sub
End Class