Imports System.IO

Module FileUtilities
    Private Function GetTempFileName(ByVal fileExtension As String) As String
        Dim file As String = Path.GetTempFileName
        g_TempFilesToDel.Add(file)
        ' Things don't get saved correctly if they don't have the right file extension
        file = Path.ChangeExtension(file, fileExtension)
        g_TempFilesToDel.Add(file)
        Return file
    End Function

    Public Function GetTempFileNameTauDemGridExt() As String
        Return GetTempFileName(TAUDEMGridExt)
    End Function

    Public Function GetTempFileNameOutputGridExt() As String
        Return GetTempFileName(OutputGridExt)
    End Function
End Module
