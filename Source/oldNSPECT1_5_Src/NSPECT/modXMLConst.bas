Attribute VB_Name = "modXMLConst"
Option Explicit

'Idea here is to store the document template for the project files in a simple string.  then just load it
Public Const g_strXMLTemplate = "<NSPECTProjectFile>" & vbNewLine & _
                                    "<PrjName></PrjName>" & vbNewLine & _
                                    "<LCGridName></LCGridName>" & vbNewLine & _
                                    "<LCGridFileName></LCGridFileName>" & vbNewLine & _
                                    "<LCGridUnits></LCGridUnits>" & vbNewLine & _
                                    "<LCGridType></LCGridType>" & vbNewLine & _
                                    "<SoilsLayerName></SoilsLayerName>" & vbNewLine & _
                                    "<SoilsFileName></SoilsFileName>" & vbNewLine & _
                                    "<SoilsAttribute></SoilsAttribute>" & vbNewLine & _
                                    "<RainFallType></RainFallType>" & vbNewLine & _
                                    "<PrecipScenario></PrecipScenario>" & vbNewLine & _
                                    "<WaterShedDelineation></WaterShedDelineation>" & vbNewLine & _
                                    "<WaterQualityStandard></WaterQualityStandard>" & vbNewLine & _
                                    "<SelectedWaterSheds></SelectedWaterSheds>" & vbNewLine & _
                                    "<LocalEffects></LocalEffects>" & vbNewLine & _
                                    "<CalcErosion></CalcErosion>" & vbNewLine & _
                                    "<ErodAttribute></ErodAttribute>" & vbNewLine & _
                                    "<RainGridBool></RainGridBool>" & vbNewLine & _
                                    "<RainGridName></RainGridName>" & vbNewLine & _
                                    "<RainGridFileName></RainGridFileName>" & vbNewLine & _
                                    "<RainConstBool></RainConstBool>" & vbNewLine & _
                                    "<RainConstValue></RainConstValue>" & vbNewLine & _
                                    "<OutPutShapefile></OutPutShapefile>" & vbNewLine & _
                                    "<OutputLayerName></OutputLayerName>" & vbNewLine & _
                                "</NSPECTProjectFile>"
'Management scenario
Public Const g_strMgmtXMLTemplate = "<MgmtScen>" & vbNewLine & _
                                        "<MgmtIgnore></MgmtIgnore>" & vbNewLine & _
                                        "<MgmtAreaLyrName></MgmtAreaLyrName>" & vbNewLine & _
                                        "<MgmtAreaFileName></MgmtAreaFileName>" & vbNewLine & _
                                        "<MgmtClassChange></MgmtClassChange>" & vbNewLine & _
                                    "</MgmtScen>"
                                    
'Public Const g_strMgmtXMLScen


