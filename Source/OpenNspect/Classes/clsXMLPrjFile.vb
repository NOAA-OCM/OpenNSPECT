'********************************************************************************************************
'File Name: clsXMLPrjFile.vb
'Description: Class for handling the project file xml
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

Imports System.Xml


Friend Class clsXMLPrjFile
    Inherits clsXMLBase
    ' *************************************************************************************
    ' *  Perot Systems Government Services
    ' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
    ' *  clsWrapperMain
    ' *************************************************************************************
    ' *  Description: XML Wrapper for use with main form's variables
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


    Private Const NODE_NAME As String = "NSPECTProjectFile"
    Private Const NODE_PRJNAME As String = "PrjName"
    Private Const NODE_PRJWORKSPACE As String = "PrjWorkspace"
    Private Const NODE_LCGridName As String = "LCGridName"
    Private Const NODE_LCGridFileName As String = "LCGridFileName"
    Private Const NODE_LCGridUnits As String = "LCGridUnits"
    Private Const NODE_LCGridType As String = "LCGridType"
    Private Const NODE_SoilsDefName As String = "SoilsDefName"
    Private Const NODE_SoilsHydFileName As String = "SoilsHydFileName"
    Private Const NODE_SoilsKFileName As String = "SoilsKFileName"
    Private Const NODE_PrecipScenario As String = "PrecipScenario"
    Private Const NODE_WaterShedDelin As String = "WaterShedDelineation"
    Private Const NODE_WaterQuality As String = "WaterQualityStandard"
    Private Const NODE_SelectedPolys As String = "SelectedPolygons"
    Private Const NODE_SelectedPolyLyrName As String = "SelectedPolyLyrName"
    Private Const NODE_SelectedPolyFileName As String = "SelectedPolyFileName"
    Private Const NODE_SelectedPolyList As String = "SelectedPolyList"
    Private Const NODE_LocalEffects As String = "LocalEffects"
    Private Const NODE_CalcErosion As String = "CalcErosion"
    '****************** Added 12/03/07 **************************
    Private Const NODE_UseOWNSDR As String = "SDRGrid"
    Private Const NODE_SDRGridFileName As String = "SDRGridFileName"
    '****************** end Add *********************************
    Private Const NODE_RainGridBool As String = "RainGridBool"
    Private Const NODE_RainGridName As String = "RainGridName"
    Private Const NODE_RainGridFileName As String = "RainGridFileName"
    Private Const NODE_RainConstBool As String = "RainConstBool"
    Private Const NODE_RainConstValue As String = "RainConstValue"

    'Variables holding value of nodes above
    Public strProjectName As String
    Public strProjectWorkspace As String
    Public strLCGridName As String
    Public strLCGridFileName As String
    Public strLCGridUnits As String
    Public strLCGridType As String
    Public strSoilsDefName As String
    Public strSoilsHydFileName As String
    Public strSoilsKFileName As String
    Public strPrecipScenario As String
    Public strWaterShedDelin As String
    Public strWaterQuality As String
    Public intSelectedPolys As Short
    Public strSelectedPolyFileName As String
    Public strSelectedPolyLyrName As String
    Public intSelectedPolyList As New Collections.Generic.List(Of Integer)
    Public intLocalEffects As Short

    'Class holders for DataGrid goodies
    Public clsMgmtScenHolder As clsXMLMgmtScenItems 'A collection of management scenarios
    Public clsPollItems As clsXMLPollutantItems 'A collection of pollutants from pollutants tab
    Public clsLUItems As clsXMLLandUseItems 'A collection of land uses
    Public clsOutputItems As clsXMLOutputItems 'A collection of outputs

    Public intCalcErosion As Short
    Public intUseOwnSDR As Short
    Public strSDRGridFileName As String
    Public intRainGridBool As Short
    Public strRainGridName As String
    Public strRainGridFileName As String
    Public intRainConstBool As Short
    Public dblRainConstValue As Double

    Public ReadOnly Property NodeName() As String
        Get
            'Retrieve the name of the element that this class wraps.

            NodeName = NODE_NAME

        End Get
    End Property


    Public Overrides Property XML() As String
        Get
            'Retrieve the XML string that this class represents. The XML returned is
            'built from the values of this class's properties.

            XML = Me.CreateNode().OuterXml

        End Get
        Set(ByVal Value As String)
            'Assign a new XML string to this class. The newly assigned XML is parsed,
            'and the class's properties are set accordingly.

            Dim dom As New XmlDocument
            Dim node As XmlNode

            If InStr(Value, ".xml") > 0 Then
                dom.Load(Value)
            Else
                dom.LoadXml(Value)
            End If

            node = dom.DocumentElement

            LoadNode(node)

        End Set
    End Property


    Public Sub SaveFile(ByRef strXML As String)
        Try
            Dim dom As New XmlDocument
            dom.LoadXml(Me.XML)

            dom.Save(strXML)

        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Public Overrides Function CreateNode(Optional ByRef Parent As XmlNode = Nothing) As XmlNode
        Try
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
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
            NodeAppendChildElement(dom, node, NODE_PRJNAME, strProjectName)
            NodeAppendChildElement(dom, node, NODE_PRJWORKSPACE, strProjectWorkspace)
            NodeAppendChildElement(dom, node, NODE_LCGridName, strLCGridName)
            NodeAppendChildElement(dom, node, NODE_LCGridFileName, strLCGridFileName)
            NodeAppendChildElement(dom, node, NODE_LCGridUnits, strLCGridUnits)
            NodeAppendChildElement(dom, node, NODE_LCGridType, strLCGridType)
            NodeAppendChildElement(dom, node, NODE_SoilsDefName, strSoilsDefName)
            NodeAppendChildElement(dom, node, NODE_SoilsHydFileName, strSoilsHydFileName)
            NodeAppendChildElement(dom, node, NODE_SoilsKFileName, strSoilsKFileName)
            'NodeAppendChildElement dom, node, NODE_RainFallType, intRainFallType
            NodeAppendChildElement(dom, node, NODE_PrecipScenario, strPrecipScenario)
            NodeAppendChildElement(dom, node, NODE_WaterShedDelin, strWaterShedDelin)
            NodeAppendChildElement(dom, node, NODE_WaterQuality, strWaterQuality)
            NodeAppendChildElement(dom, node, NODE_SelectedPolys, intSelectedPolys)
            NodeAppendChildElement(dom, node, NODE_SelectedPolyFileName, strSelectedPolyFileName)
            Dim strlist As String = ""
            If intSelectedPolyList.Count > 0 Then strlist = intSelectedPolyList(0).ToString
            For i As Integer = 1 To intSelectedPolyList.Count - 1
                strlist = strlist + "," + intSelectedPolyList(i).ToString
            Next
            NodeAppendChildElement(dom, node, NODE_SelectedPolyList, strlist)
            NodeAppendChildElement(dom, node, NODE_SelectedPolyLyrName, strSelectedPolyLyrName)
            NodeAppendChildElement(dom, node, NODE_LocalEffects, intLocalEffects)
            NodeAppendChildElement(dom, node, NODE_CalcErosion, intCalcErosion)
            NodeAppendChildElement(dom, node, NODE_UseOWNSDR, intUseOwnSDR)
            NodeAppendChildElement(dom, node, NODE_SDRGridFileName, strSDRGridFileName)
            NodeAppendChildElement(dom, node, NODE_RainGridBool, intRainGridBool)
            NodeAppendChildElement(dom, node, NODE_RainGridName, strRainGridName)
            NodeAppendChildElement(dom, node, NODE_RainGridFileName, strRainGridFileName)
            NodeAppendChildElement(dom, node, NODE_RainConstBool, intRainConstBool)
            NodeAppendChildElement(dom, node, NODE_RainConstValue, dblRainConstValue)

            'Format
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

            'Pollutants
            clsPollItems.CreateNode(node)

            'Format
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

            'Management Scenarios
            clsMgmtScenHolder.CreateNode(node)

            'Format
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
            clsLUItems.CreateNode(node)

            'Format
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))
            clsOutputItems.CreateNode(node)

            'Format
            node.AppendChild(dom.CreateTextNode(vbNewLine & vbTab))

            CreateNode = node

        Catch ex As Exception
            HandleError(ex)
            CreateNode = Nothing
        End Try
    End Function


    Public Overrides Sub LoadNode(ByRef node As XmlNode)
        Try
            'Set this class's properties based on the data found in the
            'given node.

            'Ensure that a valid node was passed in.
            If node Is Nothing Then Exit Sub

            strProjectName = GetNodeText(node, NODE_PRJNAME)
            strProjectWorkspace = GetNodeText(node, NODE_PRJWORKSPACE)
            strLCGridName = GetNodeText(node, NODE_LCGridName)
            strLCGridFileName = GetNodeText(node, NODE_LCGridFileName)
            strLCGridUnits = GetNodeText(node, NODE_LCGridUnits)
            strLCGridType = GetNodeText(node, NODE_LCGridType)
            strSoilsDefName = GetNodeText(node, NODE_SoilsDefName)
            strSoilsHydFileName = GetNodeText(node, NODE_SoilsHydFileName)
            strSoilsKFileName = GetNodeText(node, NODE_SoilsKFileName)
            strPrecipScenario = GetNodeText(node, NODE_PrecipScenario)
            strWaterShedDelin = GetNodeText(node, NODE_WaterShedDelin)
            strWaterQuality = GetNodeText(node, NODE_WaterQuality)
            intSelectedPolys = CShort(GetNodeText(node, NODE_SelectedPolys))
            strSelectedPolyFileName = GetNodeText(node, NODE_SelectedPolyFileName)
            Dim tmpstr As String() = GetNodeText(node, NODE_SelectedPolyList).Split(",")
            intSelectedPolyList.Clear()
            For i As Integer = 0 To tmpstr.Length - 1
                If tmpstr(i) <> "" Then
                    intSelectedPolyList.Add(CShort(tmpstr(i)))
                End If
            Next
            intLocalEffects = CShort(GetNodeText(node, NODE_LocalEffects))
            intCalcErosion = CShort(GetNodeText(node, NODE_CalcErosion))
            intUseOwnSDR = CShort(GetNodeText(node, NODE_UseOWNSDR, "integer"))
            strSDRGridFileName = GetNodeText(node, NODE_SDRGridFileName)
            intRainGridBool = CShort(GetNodeText(node, NODE_RainGridBool))
            strRainGridName = GetNodeText(node, NODE_RainGridName)
            strRainGridFileName = GetNodeText(node, NODE_RainGridFileName)
            intRainConstBool = CShort(GetNodeText(node, NODE_RainConstBool))
            dblRainConstValue = CDbl(GetNodeText(node, NODE_RainConstValue))

            clsMgmtScenHolder.LoadNode(node.SelectSingleNode(clsMgmtScenHolder.NodeName))
            clsPollItems.LoadNode(node.SelectSingleNode(clsPollItems.NodeName))
            clsLUItems.LoadNode(node.SelectSingleNode(clsLUItems.NodeName))
            clsOutputItems.LoadNode(node.SelectSingleNode(clsOutputItems.NodeName))
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


    Public Sub New()
        Try
            clsMgmtScenHolder = New clsXMLMgmtScenItems 'A collection of management scenarios
            clsPollItems = New clsXMLPollutantItems 'A collection of Pollutants
            clsLUItems = New clsXMLLandUseItems 'A collection of landuses
            clsOutputItems = New clsXMLOutputItems 'A collection of outputs
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub


End Class