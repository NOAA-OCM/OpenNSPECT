VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrecipType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  clsPrecipType
' *************************************************************************************
' *  Description:  This class holds the calculation strings for MUSLE based on
' *  the different Precipitation Types a user may choose
' *
' *
' *
' *  Called By:  modMUSLE
' *************************************************************************************

Private m_strCZero As String      'CZero Grid string
Private m_strCone As String       'Cone Grid string
Private m_strCTwo As String       'CTwo Grid string

'Calc string for CZero GRID in MUSLE
Public Function CZero(intType As Integer) As String
    
    Select Case strType  'intType refers to precip type 0 = TypeI, 1 = TypeIA, 2 = TypeII, 3 = Type III
        Case 0
            m_strCZero = "Con(([ip] le 0.10), 2.30550," & _
                       "Con(([ip] gt 0.10 and [ip] lt 0.20), 2.23537," & _
                       "Con(([ip] ge 0.20 and [ip] lt 0.25), 2.18219," & _
                       "Con(([ip] ge 0.25 and [ip] lt 0.30), 2.10624," & _
                       "Con(([ip] ge 0.30 and [ip] lt 0.35), 2.00303," & _
                       "Con(([ip] ge 0.35 and [ip] lt 0.40), 1.87733," & _
                       "Con(([ip] ge 0.40 and [ip] lt 0.45), 1.76312, 1.67889)))))))"
        Case 1
            m_strCZero = "Con(([ip] le 0.10), 2.03250," & _
                       "Con(([ip] gt 0.10 and [ip] lt 0.20), 1.91978," & _
                       "Con(([ip] ge 0.20 and [ip] lt 0.25), 1.83842," & _
                       "Con(([ip] ge 0.25 and [ip] lt 0.30), 1.72657, 1.63417))))"
        Case 2
            m_strCZero = "Con(([ip] le 0.10), 2.55323," & _
                       "Con(([ip] gt 0.10 and [ip] lt 0.30), 2.46532," & _
                       "Con(([ip] ge 0.30 and [ip] lt 0.35), 2.41896," & _
                       "Con(([ip] ge 0.35 and [ip] lt 0.40), 2.36409," & _
                       "Con(([ip] ge 0.40 and [ip] lt 0.45), 2.29238, 2.20282)))))"
        Case 3
            m_strCZero = "Con(([ip] le 0.10), 2.47317," & _
                       "Con(([ip] ge 0.10 and [ip] lt 0.30), 2.39628," & _
                       "Con(([ip] ge 0.30 and [ip] lt 0.35), 2.35477," & _
                       "Con(([ip] ge 0.35 and [ip] lt 0.40), 2.30726," & _
                       "Con(([ip] ge 0.40 and [ip] lt 0.45), 2.24876, 2.17772)))))"
    End Select
    
    CZero = m_strCZero
    
End Function

'Calc string for Cone GRID in MUSLE
Public Function Cone(intType As Integer) As String
    
    Select Case intType 'intType refers to precip type 0 = TypeI, 1 = TypeIA, 2 = TypeII, 3 = Type III
        Case 0
            m_strCone = "Con(([ip] le 0.10), -0.51429," & _
                        "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.50387," & _
                        "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.48488," & _
                        "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.45695," & _
                        "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.40769," & _
                        "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.32274," & _
                        "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.15644, -0.06930)))))))"
        
        Case 1
            m_strCone = "Con(([ip] le 0.10), 2.03250," & _
                        "Con(([ip] gt 0.10 and [ip] lt 0.20), 1.91978," & _
                        "Con(([ip] ge 0.20 and [ip] lt 0.25), 1.83842," & _
                        "Con(([ip] ge 0.25 and [ip] lt 0.30), 1.72657, 1.63417))))"

        Case 2
            m_strCone = "Con(([ip] le 0.10), -0.31583," & _
                        "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.28215," & _
                        "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.25543," & _
                        "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.19826, -0.09100))))"
        
        Case 3
            m_strCone = "Con(([ip] le 0.10), -0.51848," & _
                        "Con(([ip] ge 0.10 and [ip] lt 0.30), -0.51202," & _
                        "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.49735," & _
                        "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.46541," & _
                        "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.41314, -0.36803)))))"
    
    End Select
    
    Cone = m_strCone

End Function

'Calc string for CTwo GRID in MUSLE
Public Function CTwo(intType As Integer) As String

    Select Case intType 'intType refers to precip type 0 = TypeI, 1 = TypeIA, 2 = TypeII, 3 = Type III

        Case 0
            m_strCTwo = "Con(([ip] le 0.10), -0.11750," & _
                        "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.08929," & _
                        "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.06589," & _
                        "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.02835," & _
                        "Con(([ip] ge 0.30 and [ip] lt 0.35), 0.01983," & _
                        "Con(([ip] ge 0.35 and [ip] lt 0.40), 0.05754," & _
                        "Con(([ip] ge 0.40 and [ip] lt 0.45), 0.00453, 0.00000)))))))"

        Case 1
            m_strCTwo = "Con(([ip] le 0.10), -0.13748," & _
                        "Con(([ip] gt 0.10 and [ip] lt 0.20), -0.07020," & _
                        "Con(([ip] ge 0.20 and [ip] lt 0.25), -0.02597," & _
                        "Con(([ip] ge 0.25 and [ip] lt 0.30), -0.02633, -0.0))))"

        Case 2
            m_strCTwo = "Con(([ip] le 0.10), -0.16403," & _
                        "Con(([ip] gt 0.10 and [ip] lt 0.30), -0.11657," & _
                        "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.08820," & _
                        "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.05621," & _
                        "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.02281, -0.01259)))))"

        Case 3
            m_strCTwo = "Con(([ip] le 0.10), -0.17083," & _
                        "Con(([ip] ge 0.10 and [ip] lt 0.30), -0.13245," & _
                        "Con(([ip] ge 0.30 and [ip] lt 0.35), -0.11985," & _
                        "Con(([ip] ge 0.35 and [ip] lt 0.40), -0.11094," & _
                        "Con(([ip] ge 0.40 and [ip] lt 0.45), -0.11508, -0.09525)))))"

    End Select
    
    CTwo = m_strCTwo

End Function

