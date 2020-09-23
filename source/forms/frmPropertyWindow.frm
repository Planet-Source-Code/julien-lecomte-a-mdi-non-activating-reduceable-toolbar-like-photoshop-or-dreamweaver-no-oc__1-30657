VERSION 5.00
Begin VB.Form frmPropertyWindow 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   855
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPropertyWindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   57
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   237
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picCaption 
      Align           =   3  'Align Left
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   57
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
Attribute VB_Name = "frmPropertyWindow"
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

Private Sub cmdWindowAction_Click(Index As Integer)
    mnuWindowActions_Click Index
End Sub

Private Sub cmdWindowAction_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    picCaption.SetFocus
End Sub

Private Sub mnuWindowActions_Click(Index As Integer)
    Select Case Index
        Case 0
            picCaption_DblClick
        Case 1
            mdiMain.mnuWindowToolWindow(1).Checked = False
            Hide
    End Select
End Sub

Private Sub picCaption_Click()
    PopupMenu mnuClose
End Sub

Private Sub picCaption_DblClick()
    Static bHidden As Boolean
    Static lOldWidth&
    
    If bHidden Then
        Width = lOldWidth
    Else
        lOldWidth = Width
        Width = picCaption.Width * Screen.TwipsPerPixelX
    End If
    
    bHidden = Not bHidden
End Sub

Private Sub picCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case vbLeftButton
            ReleaseCapture
            SendMessageA hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End Select
End Sub

