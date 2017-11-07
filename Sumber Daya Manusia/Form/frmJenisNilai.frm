VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmJenisNilai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Jenis Nilai"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   Icon            =   "frmJenisNilai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   7350
   Begin VB.CommandButton Command2 
      Caption         =   "Tes2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tes1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   7335
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   600
         MaxLength       =   50
         TabIndex        =   11
         Top             =   5280
         Width           =   3855
      End
      Begin VB.TextBox txtKdJenisNilai 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   0
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtJenisNilai 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   5055
      End
      Begin MSDataGridLib.DataGrid dgJenisNilai 
         Height          =   3975
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   7011
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         HeadLines       =   2
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Cari"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   5280
         Width           =   285
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Nilai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Lbl 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   420
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmJenisNilai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5520
      Picture         =   "frmJenisNilai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJenisNilai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmJenisNilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command

Private Sub cmdBatal_Click()
    clear
    tampilData
    txtJenisNilai.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    If Periksa("text", txtJenisNilai, "Jenis Nilai kosong") = False Then Exit Sub
    If sp_JenisNilai("D") = False Then Exit Sub

    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"

    cmdBatal_Click
    Exit Sub
hell:
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", txtJenisNilai, "Jenis Nilai kosong") = False Then Exit Sub
    If sp_JenisNilai("A") = False Then Exit Sub

    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"

    cmdBatal_Click
    Exit Sub
hell:
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    frmCetakPenilaianPegawai.Show
End Sub

Private Sub Command2_Click()
    frmCetakPenilaianPegawaiKeDua.Show
End Sub

Private Sub dgJenisNilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisNilai.SetFocus
End Sub

Private Sub dgJenisNilai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errLoad
    If dgJenisNilai.ApproxCount = 0 Then Exit Sub
    txtKdJenisNilai.Text = dgJenisNilai.Columns(0).Value
    txtJenisNilai.Text = dgJenisNilai.Columns(1).Value
    Exit Sub
errLoad:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call tampilData
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub clear()
    txtKdJenisNilai.Text = ""
    txtJenisNilai.Text = ""
End Sub

Sub tampilData()
    On Error GoTo errTampil
    Set rs = Nothing
    strSQL = "select KdJenisNilai,JenisNilai from JenisNilai where JenisNilai like '%" & txtParameter.Text & "%' "
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgJenisNilai
        Set .DataSource = rs
        .Columns(0).DataField = rs(0).Name
        .Columns(1).DataField = rs(1).Name

        .Columns(0).Caption = "Kode"
        .Columns(0).Width = 1500
        .Columns(0).Alignment = dbgCenter
        .Columns(1).Caption = "Jenis Nilai"
        .Columns(1).Width = 4750

    End With
    Set rs = Nothing
    Exit Sub
errTampil:
    Call msubPesanError
End Sub

Private Sub txtJenisNilai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdsimpan.SetFocus
End Sub

Private Function sp_JenisNilai(f_Status As String) As Boolean
    On Error GoTo hell
    sp_JenisNilai = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisNilai", adChar, adParamInput, 2, txtKdJenisNilai.Text)
        .Parameters.Append .CreateParameter("JenisNilai", adVarChar, adParamInput, 50, Trim(txtJenisNilai.Text))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisNilai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            sp_JenisNilai = False

        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Function
hell:
    sp_JenisNilai = False
    Call msubPesanError("sp_JenisNilai")
End Function

Private Sub txtParameter_Change()
    Call tampilData
End Sub
