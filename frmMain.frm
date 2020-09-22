VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Window (Developed By Debasis Ghosh, debughosh@vsnl.net)"
   ClientHeight    =   5715
   ClientLeft      =   1950
   ClientTop       =   1575
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   960
      Width           =   1890
   End
   Begin RichTextLib.RichTextBox rtSQL 
      Height          =   2835
      Left            =   3810
      TabIndex        =   7
      Top             =   4920
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   5001
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmMain.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
      Height          =   3420
      Left            =   3810
      TabIndex        =   5
      Top             =   1410
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   6033
      _Version        =   393216
      ForeColor       =   16254213
      ForeColorFixed  =   10485760
      BackColorBkg    =   15921906
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   6300
      Left            =   195
      TabIndex        =   4
      Top             =   1425
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   11113
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   315
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   930
      Width           =   1530
   End
   Begin VB.ComboBox cmbSchema 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   195
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   945
      Width           =   5490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgRt 
      Height          =   2580
      Left            =   4770
      TabIndex        =   8
      Top             =   5070
      Visible         =   0   'False
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   4551
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Debasis Ghosh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   180
      TabIndex        =   6
      Top             =   7815
      Width           =   11595
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Schema Name And Click On Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      TabIndex        =   2
      Top             =   660
      Width           =   3555
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Node
Dim rs As New ADODB.Recordset
Dim cPb As New clsProgressBar

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLoad_Click()
    cmdLoad.Enabled = False
    tv.Nodes.Clear
    Call LoadTreeView
    cmdLoad.Enabled = True
End Sub
Private Sub Form_Load()
On Error GoTo rsError
    fg.ColWidth(0) = 300
    pb.Visible = False
    lblName.Caption = ""
    Set cPb = New clsProgressBar
    cPb.DProgressBar pb, cc3D, DRed, Standard
    Set rs = New ADODB.Recordset
    Set rs = db.OpenSchema(adSchemaSchemata, Array(Empty, Empty, Empty))
    Do Until rs.EOF
        cmbSchema.AddItem rs!SCHEMA_NAME
        rs.MoveNext
    Loop
    cmbSchema.ListIndex = 0
    rs.Close
    
'Raise Error
rsError:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        Exit Sub
    End If
End Sub
Sub LoadTreeView()
On Error GoTo rsError
    Dim c As Integer
    Set n = tv.Nodes.Add(, , "ORACLE", "" & cmbSchema.Text & " ")
    n.Expanded = True
    
    'Load Table In TreeView
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Table", "Table")
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbSchema.Text & "", Empty, "Table"))
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Table", tvwChild, "TT" & rs!TABLE_NAME, rs!TABLE_NAME)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
        
    'Load View In Treeview
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "View", "View")
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbSchema.Text & "", Empty, "View"))
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("View", tvwChild, "VV" & rs!TABLE_NAME, rs!TABLE_NAME)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
        
    'Load Procedure
    Set n = tv.Nodes.Add("ORACLE", tvwChild, "Procedure", "Procedure")
    Set rs = db.OpenSchema(adSchemaProcedures, Array(Empty, "" & cmbSchema.Text & "", Empty, Empty))
        If rs.RecordCount > 0 Then
            pb.Visible = True
            c = 1
            pb.Min = 0
            pb.Max = rs.RecordCount
            Do Until rs.EOF
                Set n = tv.Nodes.Add("Procedure", tvwChild, "PP" & rs!PROCEDURE_NAME, rs!PROCEDURE_NAME)
                pb.Value = c
                c = c + 1
                rs.MoveNext
            Loop
        End If
        pb.Visible = False
        rs.Close
'Raise Error
rsError:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        Exit Sub
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then
        rs.Close
    End If
    If db.State = 1 Then
        db.Close
    End If
End Sub
Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error GoTo rsError
    Screen.MousePointer = vbHourglass
    Dim i As Integer
    
    'Show Table
    If Node.Key = "TT" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set fg.DataSource = rs
            lblName.Caption = "Table Name: " & Node.Text & " And Record No. " & rs.RecordCount & ""
    End If
    'Show View
    If Node.Key = "VV" & Node.Text Then
        Set rs = New ADODB.Recordset
            rs.Open "Select * from " & Node.Text & "", db, adOpenDynamic, adLockBatchOptimistic
            Set fg.DataSource = rs
            lblName.Caption = "View Name: " & Node.Text & " And Record No. " & rs.RecordCount & ""
    End If
    
    'Show Procedure
    If Node.Key = "PP" & Node.Text Then
        lblName.Caption = ""
        rtSQL.Text = ""
        Set rs = db.OpenSchema(adSchemaProcedureParameters, Array(Empty, "" & cmbSchema.Text & "", "" & Node.Text & "", Empty))
        If rs.RecordCount <> 0 Then
            rtSQL.SelText = "Parameter" & vbCrLf
            rtSQL.SelText = "-------------------------------------------" & vbCrLf
            Do Until rs.EOF
                rtSQL.SelText = rs!PARAMETER_NAME & vbCrLf
                rs.MoveNext
            Loop
                rtSQL.SelText = vbCrLf
                rtSQL.SelText = vbCrLf
                rtSQL.SelText = "TEXT" & vbCrLf
                rtSQL.SelText = "-------------------------------------------" & vbCrLf
        Else
            rtSQL.SelText = "Parameter" & vbCrLf
            rtSQL.SelText = "-------------------------------------------" & vbCrLf
            rtSQL.SelText = vbCrLf
            rtSQL.SelText = vbCrLf
            rtSQL.SelText = "TEXT" & vbCrLf
            rtSQL.SelText = "-------------------------------------------" & vbCrLf
        End If
        
        'It may be wrong, Please send me the right one.
        Set rs = New ADODB.Recordset
            rs.Open "Select TEXT from ALL_SOURCE where OWNER='" & cmbSchema.Text & "' And Type = 'PROCEDURE' and NAME = '" & Trim$(Node.Text) & "'", db, adOpenDynamic, adLockBatchOptimistic
            Set fgRt.DataSource = rs
                For i = 1 To fgRt.Rows - 1
                    rtSQL.SelText = fgRt.TextMatrix(i, 1)
                Next i
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error : " & Err.Description & " ", vbCritical
        Exit Sub
    End If
End Sub
