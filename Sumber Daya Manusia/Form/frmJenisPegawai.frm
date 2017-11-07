VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmJenisPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Jenis Pegawai"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJenisPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8175
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   1560
      TabIndex        =   17
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   7920
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   19
         Top             =   5795
         Width           =   3375
      End
      Begin VB.CheckBox CheckStatusEnbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         Height          =   255
         Left            =   6360
         TabIndex        =   18
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtNamaExt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1800
         Width           =   5295
      End
      Begin VB.TextBox txtKdExt 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txtKdJenis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   2400
         MaxLength       =   3
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtJenis 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2400
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1080
         Width           =   5295
      End
      Begin MSDataGridLib.DataGrid dgJenis 
         Height          =   3405
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   6006
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
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
      Begin MSDataListLib.DataCombo dcDetailKelPeg 
         Height          =   330
         Left            =   2400
         TabIndex        =   15
         Top             =   720
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cari Jenis Pegawai"
         Height          =   210
         Left            =   2640
         TabIndex        =   20
         Top             =   5840
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detail Kelompok Pegawai"
         Height          =   210
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   2040
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nama external"
         Height          =   210
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pegawai"
         Height          =   210
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   420
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
   Begin VB.Image Image4 
      Height          =   945
      Left            =   6360
      Picture         =   "frmJenisPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJenisPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmJenisPegawai.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmJenisPegawai.frx":5A71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmJenisPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDCSource()
    strSQL = "SELECT KdDetailKelompokPegawai, DetailKelompokPegawai FROM DetailKelompokPegawai WHERE StatusEnabled = '1' order by DetailKelompokPegawai"
    Call msubDcSource(dcDetailKelPeg, rs, strSQL)
End Sub

Sub blankfield()
    On Error Resume Next
    txtKdJenis.Text = ""
    dcDetailKelPeg.BoundText = ""
    txtJenis.Text = ""
    txtKdExt.Text = ""
    txtNamaExt.Text = ""
    dcDetailKelPeg.SetFocus
    CheckStatusEnbl.Value = 1
End Sub

Sub Dag()
    strSQL = "SELECT * from V_JenisPegawai where JenisPegawai LIKE '%" & txtParameter.Text & "%' "
    Call msubRecFO(rs, strSQL)
    Set dgJenis.DataSource = rs
    dgJenis.Columns(0).Width = 1000
    dgJenis.Columns(1).Width = 0
    dgJenis.Columns(2).Width = 2000
    dgJenis.Columns(3).Width = 2000
    dgJenis.Columns(4).Width = 1500
    dgJenis.Columns(5).Width = 1500
    dgJenis.Columns(6).Width = 1300
End Sub

Private Function sp_JenisPegawai(f_Status As String) As Boolean
    sp_JenisPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisPegawai", adVarChar, adParamInput, 3, IIf(txtKdJenis.Text = "", Null, txtKdJenis.Text))
        .Parameters.Append .CreateParameter("KdDetailKelompokPegawai", adChar, adParamInput, 2, dcDetailKelPeg.BoundText)
        .Parameters.Append .CreateParameter("JenisPegawai", adVarChar, adParamInput, 50, Trim(txtJenis.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_JenisPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdCancel_Click()
    Call blankfield
    Call Dag
    Call subDCSource
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    '    On Error Resume Next
    If dgJenis.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakJenisPegawai.Show
hell:
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If Periksa("datacombo", dcDetailKelPeg, "Detail kelompok pegawai kosong") = False Then Exit Sub
    If Periksa("text", txtJenis, "Nama jenis pegawai kosong") = False Then Exit Sub
    
    strSQL = "Select * from DataPegawai where KdJenisPegawai='" & txtKdJenis & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "data tidak bisa di hapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
        Exit Sub
    Else
        If sp_JenisPegawai("D") = False Then Exit Sub
    End If
    
    
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdCancel_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errload
    
    If dcDetailKelPeg.Text <> "" Then
        If Periksa("datacombo", dcDetailKelPeg, "Detail Kelompok Pegawai Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcDetailKelPeg, "Silahkan isi nama detail kelompok pegawai ") = False Then Exit Sub
    If Periksa("text", txtJenis, "Silahkan isi nama jenis pegawai ") = False Then Exit Sub
    If sp_JenisPegawai("A") = False Then Exit Sub

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcDetailKelPeg_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtJenis.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcDetailKelPeg.Text)) = 0 Then txtJenis.SetFocus: Exit Sub
        If dcDetailKelPeg.MatchedWithList = True Then txtJenis.SetFocus: Exit Sub
        strSQL = "SELECT KdDetailKelompokPegawai, DetailKelompokPegawai FROM DetailKelompokPegawai WHERE (DetailKelompokPegawai LIKE '%" & dcDetailKelPeg.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcDetailKelPeg.BoundText = rs(0).Value
        dcDetailKelPeg.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgJenis_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenis
    WheelHook.WheelHook dgJenis
End Sub

Private Sub dgJenis_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgJenis.ApproxCount = 0 Then Exit Sub
    txtKdJenis.Text = dgJenis.Columns(0).Value
    dcDetailKelPeg.BoundText = dgJenis.Columns(1).Value
    txtJenis.Text = dgJenis.Columns(2)
    If IsNull(dgJenis.Columns(4)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgJenis.Columns(4)
    If IsNull(dgJenis.Columns(5)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgJenis.Columns(5)
    CheckStatusEnbl.Value = dgJenis.Columns(6).Value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call cmdCancel_Click

End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtNamaExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtRepDisplay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtJenis_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgJenis.SetFocus
    End Select
End Sub

Private Sub txtJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call Dag
    strCetak = " where JenisPegawai LIKE '%" & txtParameter.Text & "%'"
End Sub
