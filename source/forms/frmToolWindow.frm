VERSION 5.00
Begin VB.Form frmToolWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   555
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2430
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmToolWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   37
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   162
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTools 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   300
      TabIndex        =   3
      Top             =   75
      Width           =   1590
      Begin VB.CommandButton cmdTool 
         Height          =   315
         Index           =   4
         Left            =   1200
         Picture         =   "frmToolWindow.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Wall tool"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdTool 
         Height          =   315
         Index           =   3
         Left            =   900
         Picture         =   "frmToolWindow.frx":010E
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Image tool"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdTool 
         Height          =   315
         Index           =   2
         Left            =   600
         Picture         =   "frmToolWindow.frx":0210
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Text tool"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdTool 
         Height          =   315
         Index           =   1
         Left            =   300
         Picture         =   "frmToolWindow.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Pen tool"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdTool 
         Height          =   315
         Index           =   0
         Left            =   0
         Picture         =   "frmToolWindow.frx":0654
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ponter tool"
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox picCaption 
      Align           =   3  'Align Left
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   14
      TabIndex        =   0
      Top             =   0
      Width           =   210
      Begin VB.CommandButton cmdWindowAction 
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   30
         Width           =   180
      End
      Begin VB.CommandButton cmdWindowAction 
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   6
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   180
      End
   End
   Begin VB.Menu mnuClose 
      Caption         =   "mnuClose"
      Visible         =   0   'False
      Begin VB.Menu mnuWindowActions 
         Caption         =   "&Reduce"
         Index           =   0
      End
      Begin VB.Menu mnuWindowActions 
         Caption         =   "&Close window"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmToolWindow"
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
'
Option Explicit

Private bReduced As Boolean
    
'*******************************************************************************************
' Form routines
'*******************************************************************************************
Private Sub cmdTool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    picCaption.SetFocus
End Sub

Private Sub cmdWindowAction_Click(Index As Integer)
    mnuWindowActions_Click Index
End Sub

Private Sub cmdWindowAction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    picCaption.SetFocus
End Sub


Private Sub mnuWindowActions_Click(Index As Integer)
    Select Case Index
        Case 0: picCaption_DblClick
        Case 1
            mdiMain.mnuWindowToolWindow(0).Checked = False
            Hide
    End Select
End Sub

Private Sub picCaption_Click()
    PopupMenu mnuClose
End Sub

Private Sub picCaption_DblClick()
    Static lOldWidth&
    
    If bReduced Then
        Width = lOldWidth
    Else
        lOldWidth = Width
        Width = picCaption.Width * Screen.TwipsPerPixelX
    End If
    
    bReduced = Not bReduced
End Sub

Private Sub picCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
            ReleaseCapture
            SendMessageA hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End Select
End Sub



