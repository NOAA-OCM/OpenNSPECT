VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPathFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************
'This dll is used by the N-SPECT install script which is below:
'    function OnEnd()
'    OBJECT fixer;
'    STRING shit;
'begin
    
'    set fixer = CreateObject ( "FixPath.clsPathFix" );
    
'    If (IsObject(fixer)) Then
'        shit = fixer.FixPaths( INSTALLDIR );
'    endif;
    
'    set fixer = NOTHING;
    
'end;
'
' It will update the values of the NSPECT.mdb if the user selects and install
' directory other than C:\NSPECT
'****************************************************************************

Option Explicit
Private g_ADOConn As ADODB.Connection
' Variables used by the Error handler function - DO NOT REMOVE
Const c_sModuleFileName As String = ""


Public Function FixPaths(strInstallPath As String) As String

On Error GoTo ErrorHandler

    'Watershed Delineation
    Dim adoWSDelinRS As New ADODB.Recordset
    Dim strWSDelin As String
    
    'Soils
    Dim adoSoilsRS As New ADODB.Recordset
    Dim strSoilsRS As String
    
    'Precip
    Dim adoPrecipRS As New ADODB.Recordset
    Dim strPrecipRS As String
    
    Dim i As Integer
    
    'Strings for WSDelin Table
    Dim strDEM As String
    Dim strWS As String
    Dim strFlowDir As String
    Dim strFlowAcc As String
    Dim strDEMFill As String
    Dim strLS As String
    Dim strNibble As String
    Dim strDEM2b As String
    
    'Strings for Soils
    Dim strSoils As String
    Dim strKSoils As String
    
    'String for Precip
    Dim strPrecip As String
    
    Dim strCheck As String
    Dim strHalf As String
    Dim lngNSPECTloc As Long
  
    strWSDelin = "SELECT * FROM WSDELINEATION"
    
    If ADOConnection(strInstallPath) Then
        adoWSDelinRS.Open strWSDelin, g_ADOConn, adOpenKeyset, adLockOptimistic
        Else
            MsgBox "could not make connection"
    End If
        
        adoWSDelinRS.MoveFirst
        
        Do While Not adoWSDelinRS.EOF
        'Test on the first, that being the DEM String, if it different do it up
        strDEM = adoWSDelinRS!DEMFileName
        
        lngNSPECTloc = InStr(1, strDEM, "NSPECT", 1)
        strCheck = Left(strDEM, lngNSPECTloc + 6)
        
        If strCheck <> strInstallPath Then
        
            strHalf = Right(strDEM, Len(strDEM) - lngNSPECTloc - 6)
            strDEM = strInstallPath & strHalf
            adoWSDelinRS!DEMFileName = strDEM
            
            strWS = adoWSDelinRS!WSFileName
            strHalf = Right(strWS, Len(strWS) - lngNSPECTloc - 6)
            strWS = strInstallPath & strHalf
            adoWSDelinRS!WSFileName = strWS
            
            strFlowDir = adoWSDelinRS!FlowDirFileName
            strHalf = Right(strFlowDir, Len(strFlowDir) - lngNSPECTloc - 6)
            strFlowDir = strInstallPath & strHalf
            adoWSDelinRS!FlowDirFileName = strFlowDir
            
            strFlowAcc = adoWSDelinRS!FlowAccumFileName
            strHalf = Right(strFlowAcc, Len(strFlowAcc) - lngNSPECTloc - 6)
            strFlowAcc = strInstallPath & strHalf
            adoWSDelinRS!FlowAccumFileName = strFlowAcc
            
            strDEMFill = adoWSDelinRS!FilledDEMFileName
            strHalf = Right(strDEMFill, Len(strDEMFill) - lngNSPECTloc - 6)
            strDEMFill = strInstallPath & strHalf
            adoWSDelinRS!FilledDEMFileName = strDEMFill
            
            strLS = adoWSDelinRS!LSFileName
            strHalf = Right(strLS, Len(strLS) - lngNSPECTloc - 6)
            strLS = strInstallPath & strHalf
            adoWSDelinRS!LSFileName = strLS
            
            strNibble = adoWSDelinRS!NibbleFileName
            strHalf = Right(strNibble, Len(strNibble) - lngNSPECTloc - 6)
            strNibble = strInstallPath & strHalf
            adoWSDelinRS!NibbleFileName = strNibble
                        
            strDEM2b = adoWSDelinRS!DEM2bFileName
            strHalf = Right(strDEM2b, Len(strDEM2b) - lngNSPECTloc - 6)
            strDEM2b = strInstallPath & strHalf
            adoWSDelinRS!DEM2bFileName = strDEM2b
                       
            adoWSDelinRS.Update
            adoWSDelinRS.MoveNext
            
        Else
            adoWSDelinRS.MoveNext
        End If
        
    Loop
    
    adoWSDelinRS.Close
    'Set adoWSDelinRS = Nothing
    
    'Now do Soils
    strSoilsRS = "SELECT * FROM SOILS"
    adoSoilsRS.Open strSoilsRS, g_ADOConn, adOpenKeyset, adLockOptimistic
             
         adoSoilsRS.MoveFirst
             
    Do While Not adoSoilsRS.EOF

        strSoils = adoSoilsRS!SoilsFileName
        
        lngNSPECTloc = InStr(1, strSoils, "NSPECT", 1)
        strCheck = Left(strSoils, lngNSPECTloc + 6)
        
        If strCheck <> strInstallPath Then
        
            strHalf = Right(strSoils, Len(strSoils) - lngNSPECTloc - 6)
            strSoils = strInstallPath & strHalf
            adoSoilsRS!SoilsFileName = strSoils
    
            strKSoils = adoSoilsRS!SoilsKFileName
            strHalf = Right(strKSoils, Len(strKSoils) - lngNSPECTloc - 6)
            strKSoils = strInstallPath & strHalf
            adoSoilsRS!SoilsKFileName = strKSoils
            
            adoSoilsRS.Update
            adoSoilsRS.MoveNext
        
        Else
            adoSoilsRS.MoveNext
        End If
    Loop
    
    adoSoilsRS.Close
    'Set adoSoilsRS = Nothing
    
    'Now do Precip
    strPrecipRS = "SELECT * FROM PRECIPSCENARIO"
    
    adoPrecipRS.Open strPrecipRS, g_ADOConn, adOpenKeyset, adLockOptimistic
    
    adoPrecipRS.MoveFirst

         Do While Not adoPrecipRS.EOF
        
        strPrecip = adoPrecipRS!PrecipFileName
        
        lngNSPECTloc = InStr(1, strPrecip, "NSPECT", 1)
        strCheck = Left(strPrecip, lngNSPECTloc + 6)
        
        If strCheck <> strInstallPath Then
            
            strHalf = Right(strPrecip, Len(strPrecip) - lngNSPECTloc - 6)
            strSoils = strInstallPath & strHalf
            adoPrecipRS!PrecipFileName = strSoils
        
        
            adoPrecipRS.Update
            adoPrecipRS.MoveNext
        Else
            adoPrecipRS.MoveNext
        End If
    Loop
    
    adoPrecipRS.Close
    'Set adoPrecipRS = Nothing
    
    g_ADOConn.Close
    'Set g_ADOConn = Nothing
    
   FixPaths = "Yes"



  Exit Function
ErrorHandler:
    FixPaths = "No"
  'HandleError True, "FixPaths " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl), Err.Number, Err.Source, Err.Description, 4
End Function



'Function for connection to NSPECT.mdb: fires on dll load
Public Function ADOConnection(strInstallPath As String) As Boolean
    
    Dim strConn As String
    
On Error GoTo ErrHandler:

    Set g_ADOConn = New ADODB.Connection
        
        
    If Right(strInstallPath, 1) = "\" Then
            
        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & strInstallPath & "nspect.mdb"
    
    Else
         
        strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
           "Data Source=" & strInstallPath & "\nspect.mdb"
            
    End If
    
    g_ADOConn.Open strConn
    g_ADOConn.CursorLocation = adUseServer
    
    ADOConnection = True
Exit Function

ErrHandler:
            MsgBox "ERROR: " & Err.Number & ":  " & Err.Description
        ADOConnection = False
    
End Function

