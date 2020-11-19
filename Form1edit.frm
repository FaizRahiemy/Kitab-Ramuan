VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   12075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10320
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   10320
      TabIndex        =   16
      Top             =   240
      Width           =   1335
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
      Left            =   7200
      TabIndex        =   14
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton cmdrefresh2 
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
      Left            =   10080
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "New"
      Height          =   375
      Left            =   8760
      TabIndex        =   12
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   8760
      TabIndex        =   11
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "Save"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox TxtJenis 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox TxtMsl 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   840
      Width           =   2775
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
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1edit.frx":0000
      Top             =   1440
      Width           =   6975
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
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1edit.frx":0006
      Top             =   240
      Width           =   3735
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
      Left            =   7080
      ScaleHeight     =   4155
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   2640
      Width           =   4815
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
      Left            =   7200
      TabIndex        =   15
      Top             =   1920
      Width           =   1455
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
      Left            =   0
      TabIndex        =   6
      Top             =   1200
      Width           =   3015
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
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   1095
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
      Left            =   0
      TabIndex        =   4
      Top             =   600
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
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
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
TxtJenis.Text = ""
TxtMsl.Text = ""
TxtFarmasi.Text = ""
TxtPenanggulangan.Text = ""
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub


Private Sub CmdAbout_Click()
frmAbout.Show
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
Private Sub CmdCancel_Click()
    CmdNew.Enabled = True
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    Cmdsave.Enabled = False
    CmdCancel.Enabled = False
    CmdExit.Enabled = True
    
    TxtJenis = ""
    TxtMsl = ""
    TxtFarmasi = ""
    TxtPenanggulangan = ""
    
    TxtMsl.Enabled = False
    TxtFarmasi.Enabled = False
    TxtPenanggulangan.Enabled = False
End Sub

Private Sub CmdDelete_Click()
If TxtMsl = "" Then
    MsgBox "Silahkan Pilih Data Yang Akan Dihapus" & vbCrLf & "" _
         & "Kemudiaan Klik Delete, Terima Kasih", vbInformation, "Informasi"
    Exit Sub
End If

Dim pesan As String
pesan = MsgBox("Yakin Data Ini Akan Dihapus.....?", vbYesNo + vbQuestion, "Konfirmasi")
If pesan = vbYes Then
    Dim SQLDelete As String
    SQLDelete = "DELETE FROM Penyakit WHERE kd_msl = '" & TxtMsl & "'"
    db.Execute (SQLDelete)
    Form_Activate
    
    CmdNew.Enabled = True
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    Cmdsave.Enabled = False
    CmdCancel.Enabled = False
    CmdExit.Enabled = True
    
    TxtJenis = ""
    TxtMsl = ""
    TxtFarmasi = ""
    TxtPenanggulangan = ""
    
    TxtMsl.Enabled = False
    TxtFarmasi.Enabled = False
    TxtPenanggulangan.Enabled = False
Else
    CmdNew.Enabled = True
    CmdEdit.Enabled = True
    CmdDelete.Enabled = True
    Cmdsave.Enabled = False
    CmdCancel.Enabled = False
    CmdExit.Enabled = True
    
    TxtJenis = ""
    TxtMsl = ""
    TxtFarmasi = ""
    TxtPenanggulangan = ""
    
    TxtMsl.Enabled = False
    TxtFarmasi.Enabled = False
    TxtPenanggulangan.Enabled = False
End If
End Sub

Private Sub CmdEdit_Click()
    x = 2
    CmdEdit.Enabled = False
    CmdNew.Enabled = False
    Cmdsave.Enabled = True
    CmdCancel.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdExit.Enabled = False
    
    TxtMsl.Enabled = True
    TxtFarmasi.Enabled = True
    TxtPenanggulangan.Enabled = True
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub


Private Sub CmdNew_Click()
    TxtJenis = ""
    TxtMsl = ""
    TxtFarmasi = ""
    TxtPenanggulangan = ""
    TxtSearch = ""
    
    x = 1
    TxtJenis.Enabled = True
    TxtMsl.Enabled = True
    TxtFarmasi.Enabled = True
    TxtPenanggulangan.Enabled = True

    CmdNew.Enabled = False
    Cmdsave.Enabled = True
    CmdCancel.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdExit.Enabled = False
    
    TxtJenis.SetFocus
    
End Sub

Private Sub CmdRefresh2_Click()
TxtSearch = ""
End Sub

Private Sub CmdSave_Click()
If TxtMsl = "" Or TxtFarmasi = "" Or TxtPenanggulangan = "" Then
    MsgBox "Ada Data Yang Belum Diisi...!," & vbCrLf & "" _
           & "Mohon Data Dilengkapi Dulu", vbCritical, "Peringatan"
    Exit Sub
Else
    Dim SQLAdd As String
    Dim SQLEdit As String
    If x = 1 Then
        SQLAdd = "INSERT INTO Penyakit(jenis,kd_msl,nm_farmasi,cr_penang)values" _
               & "('" & TxtJenis & "','" & TxtMsl & "','" & TxtFarmasi & "','" & TxtPenanggulangan & "')"
        db.Execute (SQLAdd)
    ElseIf x = 2 Then
        SQLEdit = "UPDATE Penyakit SET kd_msl='" & TxtMsl & "',nm_farmasi='" & TxtFarmasi & "',cr_penang='" & TxtPenanggulangan & "' " _
                & "WHERE Jenis='" & TxtJenis & "'"
        db.Execute (SQLEdit)
    End If
End If
Form_Activate

CmdNew.Enabled = True
CmdEdit.Enabled = True
CmdDelete.Enabled = True
Cmdsave.Enabled = False
CmdCancel.Enabled = False
CmdExit.Enabled = True
    
TxtJenis = ""
TxtMsl = ""
TxtFarmasi = ""
TxtPenanggulangan = ""
    
TxtMsl.Enabled = False
TxtFarmasi.Enabled = False
TxtPenanggulangan.Enabled = False
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
    .TxtMsl.Text = MSFlexGrid1.TextMatrix(A, 1)
    .TxtFarmasi.Text = MSFlexGrid1.TextMatrix(A, 3)
    .TxtJenis.Text = MSFlexGrid1.TextMatrix(A, 2)
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
        MSFlexGrid1.TextMatrix(Baris, 5) = !video_url
        .MoveNext
        Loop
    End With
End If
End Sub

