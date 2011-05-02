Imports Microsoft.Win32

Public Class InstallationHelper
    ''' <summary>
    '''  Detects path to OpenNSPECT's data folder (installation directory) or returns null.
    ''' </summary><returns></returns>
    Public Shared Function GetInstallationDirectory() As String

        Using regKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\OpenNSPECT")
            If regKey IsNot Nothing Then
                Dim dir As String
                dir = regKey.GetValue("InstallPath")
                Return dir
            Else
                Return Nothing
            End If
        End Using
    End Function
End Class
