VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create map"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   2355
   ControlBox      =   0   'False
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   157
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHex 
      Caption         =   "Hexadecimal map"
      Height          =   240
      Left            =   75
      TabIndex        =   9
      Top             =   675
      Width           =   2190
   End
   Begin VB.TextBox txtMapName 
      Height          =   285
      Left            =   75
      MaxLength       =   64
      TabIndex        =   8
      Text            =   "Untitled"
      Top             =   300
      Width           =   2190
   End
   Begin VB.TextBox txtTiles 
      Height          =   285
      Index           =   1
      Left            =   675
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "20"
      Top             =   1485
      Width           =   465
   End
   Begin VB.TextBox txtTiles 
      Height          =   285
      Index           =   0
      Left            =   675
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "20"
      Top             =   1200
      Width           =   465
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   75
      TabIndex        =   2
      Top             =   1875
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1875
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of map :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   75
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   6
      Top             =   1500
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   1215
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number of tiles :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   4
      Top             =   975
      Width           =   1410
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    '// Check if not numeric
    If Not (IsNumeric(txtTiles(0)) And IsNumeric(txtTiles(1))) Then
        ErrMsgString ("Values must be numeric.")
        Exit Sub
    End If
    
    With tPassMapHeader
    .sName = txtMapName
    .bIsHex = CBool(chkHex = vbChecked)
    .lWidth = txtTiles(0)
    .lHeight = txtTiles(1)
    If .lWidth < 20 Or .lHeight < 20 Or .lWidth > 1000 Or .lHeight > 1000 Then
        ErrMsgString ("Values must be in between 20 and 999.")
        .lWidth = 0
        .lHeight = 0
        Exit Sub
    End If
    End With
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    lNewWindowNumber = lNewWindowNumber + 1
    txtMapName = txtMapName & " " & CStr(lNewWindowNumber)
    
    SetNumberBox txtTiles(0), True
    SetNumberBox txtTiles(1), True
End Sub
