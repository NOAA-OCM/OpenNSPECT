Attribute VB_Name = "ESRIDebugHelper"
Option Explicit

Public Sub Main()
  On Error GoTo ReregError
  Dim pCCM As IComponentCategoryManager
  Set pCCM = New ComponentCategoryManager
  Dim pCCID As New UID
  Dim pCLSID As New UID

  ' Components for registration follow

  ' CoClass: NSPECT.clsHelp
  pCLSID.Value = "{E81943A7-C2D4-4E68-9435-C4A3348DEC29}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsLCClassData
  pCLSID.Value = "{6132FB2A-1E34-4E19-ACF3-4B270300A11F}"
  ' CoClass: NSPECT.clsLCType
  pCLSID.Value = "{6EA8334D-4F18-4958-AA28-D5912FE75E0D}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsMainToolbar
  pCLSID.Value = "{F2D81016-3718-453E-BD1C-1981C04D9609}"
  pCCID.Value = "{B56A7C4A-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx CommandBars
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsMnuProject
  pCLSID.Value = "{3610E6D2-498C-4966-A2F5-93E5CB86792B}"
  ' CoClass: NSPECT.clsNewAnalysis
  pCLSID.Value = "{01A3E9C3-551A-46B1-9AB6-81C95E424B6E}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsPollutants
  pCLSID.Value = "{C74CBCA5-800D-4DEA-A06A-D3673C69E12E}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsPopUp
  pCLSID.Value = "{88BA71EA-9CFA-48FA-B695-D44347668633}"
  ' CoClass: NSPECT.clsPrecip
  pCLSID.Value = "{D2FF8F05-37ED-4811-A23A-FE6B82CA0E33}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsPrecipType
  pCLSID.Value = "{2F38FE79-27BB-456A-8143-945B612091BD}"
  ' CoClass: NSPECT.clsSoils
  pCLSID.Value = "{044A0152-A32C-40C5-B4A7-C4BCC693B572}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsToolSetup
  pCLSID.Value = "{2EE0AF71-0171-45A6-8EB8-485CFE5EBC01}"
  ' CoClass: NSPECT.clsWQStd
  pCLSID.Value = "{0DED43D4-FC74-4EAF-902F-BC6652350C5C}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsWSDelin
  pCLSID.Value = "{C2BD94A3-99D4-4B7D-B8F9-FE468C9D0D3C}"
  pCCID.Value = "{B56A7C42-83D4-11D2-A2E9-080009B6F22B}"  ' ESRI Mx Commands
  pCCM.SetupObject "C:\Projects\Workspace.vs\NSPECT\NSPECT\NSPECT.dll", pCLSID, pCCID, True

  ' CoClass: NSPECT.clsXMLCoeffTypeDef
  pCLSID.Value = "{3C022CCC-A15F-4B10-809B-3BC699F87B55}"
  ' CoClass: NSPECT.clsXMLLandUseItem
  pCLSID.Value = "{0DB19294-6298-445C-8E1A-100C5749A675}"
  ' CoClass: NSPECT.clsXMLLandUseItems
  pCLSID.Value = "{2EB42B51-73F8-4F28-B44B-D2322DCCBA58}"
  ' CoClass: NSPECT.clsXMLLUScen
  pCLSID.Value = "{4C7DADAA-863E-4C35-880F-6011552B2899}"
  ' CoClass: NSPECT.clsXMLLUScenPollItem
  pCLSID.Value = "{53542EAE-9670-4C92-B41E-C0BF92A5B469}"
  ' CoClass: NSPECT.clsXMLLUScenPollItems
  pCLSID.Value = "{D4F4E60E-0F58-439D-88F2-9C7374221C63}"
  ' CoClass: NSPECT.clsXMLMgmtScenItem
  pCLSID.Value = "{82B974EE-A2B1-413A-9C0F-8472CC0C666F}"
  ' CoClass: NSPECT.clsXMLMgmtScenItems
  pCLSID.Value = "{D8FE6959-CB77-4018-B4E9-E39EFE9F2105}"
  ' CoClass: NSPECT.clsXMLPollutantItem
  pCLSID.Value = "{02D495C0-AD18-439E-9194-0591F3B1A019}"
  ' CoClass: NSPECT.clsXMLPollutantItems
  pCLSID.Value = "{517DA047-3428-411B-957D-8E6B621A988D}"


  ' Execute the ' Debug Application
  Shell "C:\Program Files\ArcGIS\\bin\ArcMap.exe"


  Exit Sub
ReregError:
  Debug.Print "An error has occured during component re-registration"
  Err.Clear
End Sub
