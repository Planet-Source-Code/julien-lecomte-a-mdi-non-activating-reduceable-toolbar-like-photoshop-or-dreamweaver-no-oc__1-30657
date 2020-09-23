Attribute VB_Name = "modDeclare"
'*******************************************************************************************
'mdi tutorial
'*******************************************************************************************
' Julien Lecomte
' webmaster@amanitamuscaria.org
' http://www.amanitamuscaria.org
' Feel free to use, abuse or distribute. (USUS, FRUCTUS, & ABUSUS)
' If you improve it, tell me !
' Don't take credit for what you didn't create. Thanks.
'*******************************************************************************************

Option Explicit
Option Base 0

'*******************************************************************************************
' CONSTANTS
'*******************************************************************************************

Public Const DFC_BUTTON = 4
Public Const DFCS_BUTTON3STATE = &H10

Public Const GWL_HWNDPARENT = (-8)

Public Const HTCAPTION = 2
Public Const HTBOTTOMRIGHT = 17

Public Const SM_CXBORDER = 5       'Width of no-sizable borders
Public Const SM_CXDLGFRAME = 7     'Width of dialog box borders
Public Const SM_CYVSCROLL = 20     'Height of arrow in vertical scroll bar
Public Const SM_CXHSCROLL = 21     'Width of arrow in vertical scroll bar

Public Const WM_NCLBUTTONDOWN = &HA1

'*******************************************************************************************
' TYPES
'*******************************************************************************************

Public Type RECT
    Left   As Long
    Top    As Long
    Bottom As Long
    Right  As Long
End Type

'*******************************************************************************************
' USER.DLL
'*******************************************************************************************
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hDc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
