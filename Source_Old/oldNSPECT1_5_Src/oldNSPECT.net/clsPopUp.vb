Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsPopUp_NET.clsPopUp")> Public Class clsPopUp
	' *************************************************************************************
	' *  Perot Systems Government Services
	' *  Contact: Ed Dempsey - ed.dempsey@noaa.gov
	' *  clsPopup
	' *************************************************************************************
	' *  Description: Used in frmLUScen because VB does not support multiple popup menus
	' *
	' *  Called By: frmLUScen::Add/Edit Landuse scenario
	' *************************************************************************************
	' *  Subs:
	
	' *  Misc: employs clsPopUp, an API use of a popup menu.  Workaround for VB does not
	' *  support multiple popups on the same form.
	' *************************************************************************************
	
	
	Private Structure POINT
		Dim X As Integer
		Dim Y As Integer
	End Structure
	
	Private Const MF_ENABLED As Integer = &H0
	Private Const MF_SEPARATOR As Integer = &H800
	Private Const MF_STRING As Integer = &H0
	Private Const TPM_RIGHTBUTTON As Integer = &H2
	Private Const TPM_LEFTALIGN As Integer = &H0
	Private Const TPM_NONOTIFY As Integer = &H80
	Private Const TPM_RETURNCMD As Integer = &H100
	Private Declare Function CreatePopupMenu Lib "user32" () As Integer
	Private Declare Function AppendMenu Lib "user32"  Alias "AppendMenuA"(ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal wIDNewItem As Integer, ByVal sCaption As String) As Integer
	Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Integer, ByVal wFlags As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nReserved As Integer, ByVal hwnd As Integer, ByRef nIgnored As Integer) As Integer
	Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Integer) As Integer
	'UPGRADE_WARNING: Structure POINT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINT) As Integer
	Private Declare Function GetForegroundWindow Lib "user32" () As Integer
	'
	'UPGRADE_WARNING: ParamArray Param was changed from ByRef to ByVal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="93C6A0DC-8C99-429A-8696-35FC4DCEFCCC"'
	Public Function Popup(ParamArray ByVal Param() As Object) As Integer
		Dim iMenu As Integer
		Dim hMenu As Integer
		Dim nMenus As Integer
		Dim p As POINT
		
		' get the current cursor pos in screen coordinates
		GetCursorPos(p)
		
		' create an empty popup menu
		hMenu = CreatePopupMenu()
		
		' determine # of strings in paramarray
		nMenus = 1 + UBound(Param)
		
		' put each string in the menu
		For iMenu = 1 To nMenus
			' the AppendMenu function has been superseeded by the InsertMenuItem
			' function, but it is a bit easier to use.
			'UPGRADE_WARNING: Couldn't resolve default property of object Param(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Trim(CStr(Param(iMenu - 1))) = "-" Then
				' if the parameter is a single dash, a separator is drawn
				AppendMenu(hMenu, MF_SEPARATOR, iMenu, "")
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object Param(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				AppendMenu(hMenu, MF_STRING + MF_ENABLED, iMenu, CStr(Param(iMenu - 1)))
			End If
		Next iMenu
		
		' show the menu at the current cursor location;
		' the flags make the menu aligned to the right (!); enable the right button to select
		' an item; prohibit the menu from sending messages and make it return the index of
		' the selected item.
		' the TrackPopupMenu function returns when the user selected a menu item or cancelled
		' the window handle used here may be any window handle from your application
		' the return value is the (1-based) index of the menu item or 0 in case of cancelling
		iMenu = TrackPopupMenu(hMenu, TPM_RIGHTBUTTON + TPM_LEFTALIGN + TPM_NONOTIFY + TPM_RETURNCMD, p.X, p.Y, 0, GetForegroundWindow(), 0)
		
		' release and destroy the menu (for sanity)
		DestroyMenu(hMenu)
		
		' return the selected menu item's index
		Popup = iMenu
		
	End Function
End Class