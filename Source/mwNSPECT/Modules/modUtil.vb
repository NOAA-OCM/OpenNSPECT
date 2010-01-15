Imports System.Data.OleDb
Imports System.Windows.Forms

Module modUtil
    Public g_nspectPath As String
    Public g_nspectDocPath As String

    'Database Variables
    Public g_DBConn As OleDbConnection 'Connection
    Public g_strConn As String 'Connection String
    Public g_boolConnected As Boolean 'Bool: connected

    Public g_boolAddCoeff As Boolean
    Public g_boolCopyCoeff As Boolean 'True: called frmPollutants, False: called frmNewPollutants
    Public g_boolAgree As Boolean 'True: use the Agree Function on Streams.
    Public g_boolHydCorr As Boolean 'True: Hyrdologically Correct DEM, no fill needed
    Public g_boolNewWShed As Boolean 'True: New WaterShed form called from frmPrj


    'WqStd

    'Agree DEM Stuff
    Public g_boolParams As Boolean 'Flag to indicate Agree params have been entered


    'Project Form Variables
    Public g_strPrjFileName As String 'Project file name

    'Management Scenario variables::frmPrjCalc
    Public g_strLUScenFileName As String 'Management scenario file name
    Public g_intManScenRow As String 'Management scenario ROW number

    'Pollutant Coefficient variable::frmPrjCalc
    Public g_intCoeffRow As Short 'Coeff Row Number
    Public g_strCoeffCalc As String 'if the Calc option is chosen, hold results in string


    Const c_sModuleFileName As String = "modUtil.vb"
    Private m_ParentHWND As Integer ' Set this to get correct parenting of Error handler forms



    'Function for connection to NSPECT.mdb: fires on dll load
    Public Sub DBConnection()
        Try
            If Not g_boolConnected Then
                'TODO: check for location of file and prompt if not found
                g_strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & g_nspectPath & "\nspect.mdb"

                g_DBConn = New OleDbConnection(g_strConn)

                g_DBConn.Open()

                g_boolConnected = True

            End If
            Exit Sub
        Catch ex As Exception
            MsgBox(Err.Number & Err.Description & " Error connecting to database, please check NSPECTDAT enviornment variable.  Current value of NSPECTDAT: " & g_strConn, MsgBoxStyle.Critical, "Error Connecting")

        End Try
    End Sub



    Public Sub InitComboBox(ByRef cbo As System.Windows.Forms.ComboBox, ByRef strName As String)
        Try
            'Loads the variety of comboboxes throught the project using combobox and name of table
            Dim rsNamesCmd As OleDbCommand
            Dim rsNames As OleDbDataReader
            Dim strSelectStatement As String

            strSelectStatement = "SELECT NAME FROM " & strName & " ORDER BY NAME ASC"

            'Check thrown in to make sure g_ADOconn is something, in v9.1 we started having problems.
            If Not g_boolConnected Then
                DBConnection()
            End If

            rsNamesCmd = New OleDbCommand(strSelectStatement, g_DBConn)

            rsNames = rsNamesCmd.ExecuteReader()

            If rsNames.HasRows Then
                With cbo
                    Do While rsNames.Read()
                        .Items.Add(rsNames.Item("Name"))
                    Loop
                End With

                cbo.SelectedIndex = 0
            Else
                MsgBox("Warning.  There are no records remaining.  Please add a new one.", MsgBoxStyle.Critical, "Recordset Empty")
                Exit Sub
            End If

            'Cleanup
            rsNames.Close()
        Catch ex As Exception
            HandleError(True, "InitComboBox " & c_sModuleFileName & " " & GetErrorLineNumberString(Erl()), Err.Number, Err.Source, Err.Description, 1, m_ParentHWND)
        End Try
    End Sub


    'Returns a filename given for example C:\temp\dataset returns dataset
    Public Function SplitFileName(ByRef sWholeName As String) As String
        SplitFileName = ""
        Try
            Dim pos As Short
            Dim sT As Object
            Dim sName As String
            pos = InStrRev(sWholeName, "\")
            If pos > 0 Then
                sT = Mid(sWholeName, 1, pos - 1)
                If pos = Len(sWholeName) Then
                    Exit Function
                End If
                sName = Mid(sWholeName, pos + 1, Len(sWholeName) - Len(sT))
                pos = InStr(sName, ".")
                If pos > 0 Then
                    SplitFileName = Mid(sName, 1, pos - 1)
                Else
                    SplitFileName = sName
                End If
            End If
        Catch ex As Exception
            MsgBox("Workspace Split:" & Err.Description)
        End Try
    End Function



End Module
