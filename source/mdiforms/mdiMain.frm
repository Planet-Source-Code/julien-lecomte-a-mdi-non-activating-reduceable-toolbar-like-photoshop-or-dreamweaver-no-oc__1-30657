VERSION 5.00
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "mdi tutorial"
   ClientHeight    =   5265
   ClientLeft      =   165
   ClientTop       =   705
   ClientWidth     =   7500
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowTile 
         Caption         =   "Tile &Horizontally"
         Index           =   0
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "Tile &Vertically"
         Index           =   1
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Cascade"
         Index           =   2
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Arrange Icons"
         Index           =   3
      End
      Begin VB.Menu mnuWindowToolWindowSepar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowToolWindow 
         Caption         =   "Tool window"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuWindowToolWindow 
         Caption         =   "Property window"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnuWindowListSepar 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuWindowList 
         Caption         =   "list"
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'NOTES:
' frmPropertyWindow & frmToolWindow are the non activating mdi child tool windows.
' Because they must be childs, they appear in the windowlist menu, that is why the menu
' has been altered, and the menu is now created by the MakeWindowList procedure.

Option Explicit
Option Base 0

Private Const OPT_FULLY = 0
Private Const OPT_PARTIAL = 1
Private Const OPT_NONE = 2

'// Change this value to another OPT value to see the different effects if the the
'// mdi parent is resized and the tool box becomes invisible.
Private Const Option_byToolWindowResize = OPT_FULLY

Private Sub MDIForm_Load()
    '// Load toolbars
    Load frmToolWindow
    Load frmPropertyWindow

    SetWindowMDIToolWindow frmToolWindow, mdiMain
    SetWindowMDIToolWindow frmPropertyWindow, mdiMain

    frmToolWindow.Show
    frmPropertyWindow.Show
End Sub

Private Sub MDIForm_Resize()
    ToolWindow_RePos frmToolWindow
    ToolWindow_RePos frmPropertyWindow
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim I&
    
    '// Unload all new menus
    For I = mnuWindowList.UBound To mnuWindowList.LBound Step -1
        If I <> 0 Then Unload mnuWindowList(I)
    Next
    
    '// Unloads all child windows, toolbars included
    Unload frmToolWindow
    Unload frmPropertyWindow
End Sub

Private Sub mnuFileNew_Click()
    Screen.MousePointer = vbHourglass
    Dim objChildForm As New frmMap
    objChildForm.Show
    objChildForm.Form_Activate
    Set objChildForm = Nothing
    Screen.MousePointer = vbNormal

End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Public Sub MakeWindowList()
    Dim I&
    Dim frmChild As Form
    
    '// Menu 0 is already kept hidden to be able to load & unload new menus
    mnuWindowListSepar.Visible = False
    
    For I = mnuWindowList.UBound To mnuWindowList.LBound Step -1
        If I Then Unload mnuWindowList(I)
    Next
    
    I = 1
    For Each frmChild In Forms
        If frmChild.Name <> frmToolWindow.Name And _
           frmChild.Name <> frmPropertyWindow.Name And _
           frmChild.Name <> mdiMain.Name Then
            mnuWindowListSepar.Visible = True

            '// Load window in menu
            Load mnuWindowList(I)
            mnuWindowList(I).Caption = frmChild.Caption
            mnuWindowList(I).Visible = True
            mnuWindowList(I).Tag = frmChild.hWnd
                        
            '// If is active window then activate as active window, otherwise find topmost window
            mnuWindowList(I).Checked = CBool(mdiMain.ActiveForm.hWnd = frmChild.hWnd)
            
            I = I + 1
        End If
    Next
    
End Sub


Private Sub mnuWindowList_Click(Index As Integer)
    Dim frmChild As Form
    Dim hWndToBeActivated&
    
    hWndToBeActivated = mnuWindowList(Index).Tag
        
    For Each frmChild In Forms
        If frmChild.Name <> frmToolWindow.Name And _
           frmChild.Name <> frmPropertyWindow.Name And _
           frmChild.Name <> mdiMain.Name Then
           
            If frmChild.hWnd = hWndToBeActivated Then
                frmChild.SetFocus
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Sub mnuWindowTile_Click(Index As Integer)
    Select Case Index
        Case 0: mdiMain.Arrange (vbTileHorizontal)
        Case 1: mdiMain.Arrange (vbTileVertical)
        Case 2: mdiMain.Arrange (vbCascade)
        Case 3: mdiMain.Arrange (vbArrangeIcons)
    End Select
End Sub

Private Sub mnuWindowToolWindow_Click(Index As Integer)
    Dim objForm As Form
    
    Select Case Index
        Case 0: Set objForm = frmToolWindow
        Case 1: Set objForm = frmPropertyWindow
    End Select
    mnuWindowToolWindow(Index).Checked = Not mnuWindowToolWindow(Index).Checked
    If mnuWindowToolWindow(Index).Checked Then
        objForm.Show
        ToolWindow_RePos objForm
    Else
        objForm.Hide
    End If
On Error Resume Next
    objForm.picCaption.SetFocus
    Set objForm = Nothing
End Sub

Private Sub ToolWindow_RePos(xForm As Form)
    Dim lXMdi&
    
    If Not xForm.Visible Then Exit Sub
    If WindowState = vbMinimized Then Exit Sub
    If Option_byToolWindowResize = OPT_NONE Then Exit Sub
    
    lXMdi = (GetSystemMetrics(SM_CXDLGFRAME) + GetSystemMetrics(SM_CXBORDER)) * Screen.TwipsPerPixelX
    
    Select Case Option_byToolWindowResize
        Case OPT_PARTIAL
            If xForm.Left < lXMdi Then
                xForm.Left = lXMdi
            ElseIf xForm.Left >= mdiMain.ScaleWidth Then
                xForm.Left = mdiMain.ScaleWidth - xForm.Width
            End If
            
            If xForm.Top < lXMdi Then
                xForm.Top = lXMdi
            ElseIf xForm.Top >= mdiMain.ScaleHeight Then
                xForm.Top = mdiMain.ScaleHeight - xForm.Height
            End If
            
        Case OPT_FULLY
            If xForm.Left < lXMdi Then
                xForm.Left = lXMdi
            ElseIf xForm.Left + xForm.Width >= mdiMain.ScaleWidth Then
                xForm.Left = mdiMain.ScaleWidth - xForm.Width
            End If
            
            If xForm.Top < lXMdi Then
                xForm.Top = lXMdi
            ElseIf xForm.Top + xForm.Height >= mdiMain.ScaleHeight Then
                xForm.Top = mdiMain.ScaleHeight - xForm.Height
            End If
            
    End Select
End Sub
