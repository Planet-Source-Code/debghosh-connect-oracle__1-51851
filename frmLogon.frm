VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon"
   ClientHeight    =   3735
   ClientLeft      =   2625
   ClientTop       =   2550
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6150
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   465
      Left            =   4140
      TabIndex        =   7
      Top             =   2970
      Width           =   1590
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   465
      Left            =   2475
      TabIndex        =   6
      Top             =   2970
      Width           =   1605
   End
   Begin VB.TextBox txtDatabase 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2000
      TabIndex        =   5
      Top             =   1860
      Width           =   3900
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1185
      Width           =   3900
   End
   Begin VB.TextBox txtUserId 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2000
      TabIndex        =   1
      Top             =   525
      Width           =   3900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter USER ID, Password and Database Name."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   915
      TabIndex        =   8
      Top             =   210
      Width           =   4440
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   6015
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter Database"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   4
      Top             =   1995
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter PasssWord"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Shape Shape1 
      Height          =   3480
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   135
      Width           =   5925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter USER ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      TabIndex        =   0
      Top             =   630
      Width           =   1335
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOk_Click()
    OracleConnect
End Sub
Private Sub Form_Load()
    Set txtSubclass.txt = txtPassword
    Hook
End Sub
Private Sub Form_Unload(Cancel As Integer)
    UnHook
End Sub
