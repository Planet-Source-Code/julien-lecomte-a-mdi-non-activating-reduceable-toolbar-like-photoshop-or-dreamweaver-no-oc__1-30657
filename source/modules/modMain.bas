Attribute VB_Name = "modMain"
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
' NOTES in mdiMain
Option Explicit

Public lNewWindowNumber As Long       '// Untitled map number, incremented every time

'*******************************************************************************************
' Main
'*******************************************************************************************
Public Sub Main()
    mdiMain.Show
End Sub

Public Sub SetWindowMDIToolWindow(objWindow As Form, objMDIForm As MDIForm)
    Dim R As RECT
    
    SetRect R, objWindow.picCaption.ScaleWidth, 0&, objWindow.ScaleWidth, objWindow.ScaleHeight
    DrawFrameControl objWindow.hDc, R, DFC_BUTTON, DFCS_BUTTON3STATE
    SetParent objWindow.hWnd, objMDIForm.hWnd
    SetWindowLong objWindow.hWnd, GWL_HWNDPARENT, objMDIForm.hWnd
    
    '// For non mdi toolwindows, use this instead
    'SetWindowPos objWindow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
