VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFarmasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kitab Ramuan XI-I"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12525
   Icon            =   "Form1.frm.bkp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frm.bkp.frx":058A
   ScaleHeight     =   7170
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdrefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11150
      TabIndex        =   11
      Top             =   1030
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4335
      Left            =   7560
      TabIndex        =   10
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7646
      _Version        =   393216
      BackColorBkg    =   16777215
      ScrollBars      =   2
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8520
      TabIndex        =   9
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   6840
      Width           =   1215
   End
   Begin VB.TextBox TxtFarmasi 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "Form1.frm.bkp.frx":EB68
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox TxtPenanggulangan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form1.frm.bkp.frx":EB6E
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label LblJenis 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Masalah/farmasi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label LblMsl 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Penyebab"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Penanggulangan 
      BackStyle       =   0  'Transparent
      Caption         =   "Penanggulangan/manfaat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   3015
   End
End
Attribute VB_Name = "FrmFarmasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim panjang As Integer
'FIXIT: Declare 'hapus' with an early-bound data type                                      FixIT90210ae-R1672-R1B8ZE
Dim hapus, keyword As String
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Dim x As Integer
Dim RSCari As New ADODB.Recordset
Sub bersih()
LblJenis.Caption = ""
LblMsl.Caption = ""
TxtFarmasi.Text = ""
TxtPenanggulangan.Text = ""
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
frmAbout.Show
End Sub

Private Sub Cmdrefresh_Click()
TxtSearch = ""
End Sub

Private Sub Form_Load()
Call bersih
Me.Top = 2500
Me.Left = 3750
End Sub
Private Sub Form_Activate()
Koneksi
TxtSearch.SetFocus
db.CursorLocation = adUseClient
MSFlexGrid1.Refresh
TampilGrid
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub
Sub AktifGrid()
With MSFlexGrid1
    .Cols = 5
    .RowHeightMin = 300
        
    .Col = 0
    .Row = 0
    .Text = "NO"
    .CellFontBold = True
    .ColWidth(0) = 400
    .AllowUserResizing = flexResizeColumns
    .CellAlignment = flexAlignCenterCenter
      
    .Col = 1
    .Row = 0
    .Text = "MASALAH/FARMASI"
    .CellFontBold = True
    .ColWidth(1) = 2200
    .AllowUserResizing = flexResizeColumns
    .CellAlignment = flexAlignCenterCenter
        
    .Col = 2
    .Row = 0
    .Text = "JENIS"
    .CellFontBold = True
    .ColWidth(2) = 1800
    .AllowUserResizing = flexResizeColumns
    .CellAlignment = flexAlignCenterCenter
    
    .Col = 3
    .Row = 0
    .Text = "PENYEBAB"
    .CellFontBold = True
    .ColWidth(3) = 3000
    .AllowUserResizing = flexResizeColumns
    .CellAlignment = flexAlignCenterCenter
    
    .Col = 4
    .Row = 0
    .Text = "PENANGGULANGAN"
    .CellFontBold = True
    .ColWidth(4) = 5000
    .AllowUserResizing = flexResizeColumns
    .CellAlignment = flexAlignCenterCenter
End With
End Sub
Sub TampilGrid()
Dim Baris As String
MSFlexGrid1.Clear
Call AktifGrid
    
MSFlexGrid1.Rows = 2
Baris = 0
Call Koneksi
rs.Open "SELECT * FROM Penyakit ORDER BY kd_msl ASC", db, adOpenDynamic, adLockOptimistic

If rs.BOF Then
    Exit Sub
Else
    With rs
        .MoveFirst
        Do While Not .EOF
        On Error Resume Next
        Baris = Baris + 1
        MSFlexGrid1.Rows = Baris + 1
        MSFlexGrid1.TextMatrix(Baris, 0) = Baris
        MSFlexGrid1.TextMatrix(Baris, 1) = !kd_msl
        MSFlexGrid1.TextMatrix(Baris, 2) = !jenis
        MSFlexGrid1.TextMatrix(Baris, 3) = !nm_farmasi
        MSFlexGrid1.TextMatrix(Baris, 4) = !cr_penang
        .MoveNext
    Loop
    End With
End If
End Sub
Private Sub MSFlexGrid1_DblClick()
Dim A As Integer
A = MSFlexGrid1.Row
If MSFlexGrid1.Rows <> 1 Then
    With Me
    .LblMsl.Caption = MSFlexGrid1.TextMatrix(A, 1)
    .TxtFarmasi.Text = MSFlexGrid1.TextMatrix(A, 3)
    .LblJenis.Caption = MSFlexGrid1.TextMatrix(A, 2)
    .TxtPenanggulangan.Text = MSFlexGrid1.TextMatrix(A, 4)
    End With
Else
    Exit Sub
End If
End Sub
Private Sub TxtSearch_Change()
    Set RSCari = New ADODB.Recordset
    RSCari.Open "SELECT * FROM Penyakit WHERE kd_msl like '%" & TxtSearch.Text & "%'", db, 1, 1
    TampilGrid1
End Sub
Sub TampilGrid1()
Dim Baris As String
MSFlexGrid1.Clear
Call AktifGrid
    
MSFlexGrid1.Rows = 2
Baris = 0

If RSCari.BOF Then
    Exit Sub
Else
    With RSCari
        On Error Resume Next
        .MoveFirst
        Do While Not .EOF
        Baris = Baris + 1
        MSFlexGrid1.Rows = Baris + 1
        MSFlexGrid1.TextMatrix(Baris, 0) = Baris
        MSFlexGrid1.TextMatrix(Baris, 1) = !kd_msl
        MSFlexGrid1.TextMatrix(Baris, 2) = !jenis
        MSFlexGrid1.TextMatrix(Baris, 3) = !nm_farmasi
        MSFlexGrid1.TextMatrix(Baris, 4) = !cr_penang
        .MoveNext
        Loop
    End With
End If
End Sub
