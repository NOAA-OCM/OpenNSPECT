Attribute VB_Name = "modFrmPrjWinPosition"
' *************************************************************************************
' *  Perot Systems Government Services
' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
' *  modfrmPrjWinPosition
' *************************************************************************************
' *  Description:  Keeps the frmPrj in front of the ArcMap window
' *
' *  Subs/Functions:
' *
' *  Called By:  frmPrj
' *************************************************************************************




Option Explicit

' WINDOW POSITIONING
Declare Function SetWindowPos Lib "user32" ( _
ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const FLAGS As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Public Sub SetWin_NOTOPMOST(hWnd As Long)
    SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Public Sub SetWin_TOPMOST(hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub


