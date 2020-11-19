VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form FrmFarmasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kitab Ramuan XI-I"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0CCA
   ScaleHeight     =   8505
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit Database"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   8160
      Width           =   1215
   End
   Begin MCI.MMControl M1 
      Height          =   330
      Left            =   2880
      TabIndex        =   13
      Top             =   8160
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   582
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
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
      Height          =   255
      Left            =   8040
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.PictureBox MSFlexGrid1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   7320
      ScaleHeight     =   4155
      ScaleWidth      =   4755
      TabIndex        =   10
      Top             =   2520
      Width           =   4815
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
      Left            =   7320
      TabIndex        =   9
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton CmdAbout 
      Caption         =   "About"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   8160
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
      Text            =   "Form1.frx":2869C
      Top             =   1320
      Width           =   3735
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
      Height          =   5415
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Form1.frx":286A2
      Top             =   2520
      Width           =   6975
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10200
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   9960
      Picture         =   "Form1.frx":286A8
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   11520
      Picture         =   "Form1.frx":28D63
      Top             =   1200
      Width           =   660
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
      Left            =   7320
      TabIndex        =   12
      Top             =   1320
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


Private Sub CmdAbout_Click()
frmAbout.Show
End Sub

Private Sub CmdEdit_Click()
MsgBox ("Fitur Ini Belum Tersedia")
End Sub

Private Sub CmdRefresh_Click()
TxtSearch = ""
End Sub

Private Sub Form_Load()
Call bersih
Me.Top = 2500
Me.Left = 3750

End Sub
Private Sub Form_Activate()
Koneksi

db.CursorLocation = adUseClient
MSFlexGrid1.Refresh
TampilGrid
End Sub
Private Sub Form_Initialize()
    InitCommonControls

End Sub
Sub AktifGrid()
With MSFlexGrid1
    .Cols = 6
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
        
    .Col = 5
    .Row = 0
    .Text = "VIDEO URL"
    .CellFontBold = True
    .ColWidth(5) = 5000
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
        MSFlexGrid1.TextMatrix(Baris, 5) = !video_url
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
    .M1.FileName = MSFlexGrid1.TextMatrix(A, 5)
    .Label4.Caption = MSFlexGrid1.TextMatrix(A, 1)
    End With
    Image2.Picture = LoadPicture(App.Path + "\cloud kosong.gif")
    Label4.Visible = True
    If LblMsl.Caption = "Batuk" Then
    M1.FileName = App.Path + "\video\batuk.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    ElseIf LblMsl.Caption = "Radang Tenggorokan" Then
    M1.FileName = App.Path + "\video\Radang Tenggorokan.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    ElseIf LblMsl.Caption = "Jerawat" Then
    M1.FileName = App.Path + "\video\jerawat.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    ElseIf LblMsl.Caption = "Sakit Gigi" Then
    M1.FileName = App.Path + "\video\Sakit Gigi.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    ElseIf LblMsl.Caption = "Masuk Angin" Then
    M1.FileName = App.Path + "\video\masuk angin.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    ElseIf LblMsl.Caption = "Pilek" Then
    M1.FileName = App.Path + "\video\pilek.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    ElseIf LblMsl.Caption = "Ketombe" Then
    M1.FileName = App.Path + "\video\ketombe.avi"
    M1.Command = "close"
    M1.Command = "open"
    M1.Command = "play"
    End If
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
        MSFlexGrid1.TextMatrix(Baris, 5) = !video_url
        .MoveNext
        Loop
    End With
End If
End Sub
