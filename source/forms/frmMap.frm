VERSION 5.00
Begin VB.Form frmMap 
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5085
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   Begin VB.PictureBox picPicture 
      BackColor       =   &H00C0C0FF&
      Height          =   2190
      Left            =   225
      ScaleHeight     =   2130
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   300
      Width           =   2415
   End
   Begin VB.HScrollBar slbH 
      Height          =   240
      Left            =   0
      Max             =   1
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3375
      Width           =   4665
   End
   Begin VB.VScrollBar slbV 
      Height          =   3315
      Left            =   4800
      Max             =   1
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblSizeBox 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   255
      Index           =   0
      Left            =   4725
      MousePointer    =   8  'Size NW SE
      TabIndex        =   3
      Top             =   3300
      Width           =   255
   End
   Begin VB.Label lblSizeBox 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000016&
      Height          =   255
      Index           =   1
      Left            =   4800
      MousePointer    =   6  'Size NE SW
      TabIndex        =   2
      Top             =   3375
      Width           =   255
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
'

'*******************************************************************************************
' Form
'*******************************************************************************************
Private Sub Form_Load()
    lNewWindowNumber = lNewWindowNumber + 1
    Caption = "Untitled " & lNewWindowNumber
End Sub

Public Sub Form_Activate()
    mdiMain.MakeWindowList
End Sub

Public Sub Form_Resize()
    Dim lX&, lY&
    Dim hscrl&, vscrl&
    
    If WindowState = vbMinimized Then Exit Sub
    
    '// Get system metrics
    hscrl = GetSystemMetrics(SM_CXHSCROLL)
    vscrl = GetSystemMetrics(SM_CYVSCROLL)
    lX = ScaleWidth - vscrl
    lY = ScaleHeight - hscrl
    
    '// Move scrollbars
    slbH.Move 0&, lY, lX, hscrl
    slbV.Move lX, 0&, vscrl, lY
    
    '// If there is a picture box, then move it full screen
    picPicture.Move 0&, 0&, lX, lY
    
    '// Set scrollbar values here
    
    '// Move or hide resize handlers
    lblSizeBox(0).Visible = Not CBool(WindowState = vbMaximized)
    lblSizeBox(1).Visible = Not CBool(WindowState = vbMaximized)
    lblSizeBox(0).Width = hscrl
    lblSizeBox(1).Width = hscrl
    lblSizeBox(0).Height = vscrl
    lblSizeBox(1).Height = vscrl
    lblSizeBox(0).Move ScaleWidth - lblSizeBox(0).Width, ScaleHeight - lblSizeBox(0).Height
    lblSizeBox(1).Move ScaleWidth - lblSizeBox(1).Width - 1, ScaleHeight - lblSizeBox(1).Height - 1
End Sub

Private Sub Form_Terminate()
    mdiMain.MakeWindowList
End Sub

'*******************************************************************************************
' Window elements
'*******************************************************************************************
Private Sub lblSizeBox_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    '// Resize capture
    ReleaseCapture
    SendMessageA hWnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, 0&
End Sub
