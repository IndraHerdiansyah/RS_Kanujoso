VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmJenjangJabatan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Jenjang Jabatan "
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmJenjangJabatan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   6975
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   960
      TabIndex        =   15
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox txtKdExt 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1800
      MaxLength       =   15
      TabIndex        =   12
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtNamaExt 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1800
      MaxLength       =   30
      TabIndex        =   11
      Top             =   2640
      Width           =   5055
   End
   Begin VB.CheckBox chkStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Status Aktif"
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtKode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   7
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtNama 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   6960
      Width           =   1095
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
   Begin MSDataGridLib.DataGrid dgTit 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Kode External"
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nama external"
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Jenjang Jabatan"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Kode"
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   420
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmJenjangJabatan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5280
      Picture         =   "frmJenjangJabatan.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJenjangJabatan.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmJenjangJabatan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoCommand As New ADODB.Command
Dim vbMsgboxRslt As VbMsgBoxResult

Private Sub cmdBatal_Click()
    clear
    tampilData
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    If dgTit.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmCetakJenjangJabatan.Show
hell:
End Sub

Private Sub cmdHapus_Click()
'    On Error GoTo xxx
'    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
'    dbConn.Execute "DELETE JenjangJabatan WHERE KdJenjang = '" & txtkode.Text & "'"
'    tampilData
'    clear
'    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
'    Exit Sub
'xxx:
'    MsgBox "Data digunakan, tidak dapat dihapus..", vbOKOnly, "Validasi"
    
On Error GoTo Errload
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    If Periksa("text", txtNama, "Pilih Data yang akan dihapus") = False Then Exit Sub
    strSQL = "Select * from V_M_DataPegawaiNew where KdJenjang='" & txtkode & "'"
    Call msubRecFO(rs, strSQL)
    If rs.EOF = False Then
        MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
        Exit Sub
    Else
        dbConn.Execute "DELETE JenjangJabatan WHERE KdJenjang = '" & txtkode.Text & "'"
        tampilData
        clear
        MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
        Exit Sub
    End If
Errload:
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell

    If Periksa("text", txtNama, "Silahkan isi nama jenjang jabatan fungsional ") = False Then Exit Sub

    With adoCommand
        .ActiveConnection = dbConn
        .CommandType = adCmdStoredProc
        .CommandText = "AUD_JenjangJabatan"

        If txtkode.Text = "" Then
            .Parameters.Append .CreateParameter("KdJenjang", adChar, adParamInput, 5, Null)
        Else
            .Parameters.Append .CreateParameter("KdJenjang", adChar, adParamInput, 5, txtkode.Text)
        End If

        .Parameters.Append .CreateParameter("NamaJenjangJabatan", adVarChar, adParamInput, 50, Trim(txtNama.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatus.Value)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, "A")

        If Not IsNull(.Parameters("KdJenjang").Value) Then
            MsgBox "Data berhasil di update ..", vbInformation, "Informasi"
            .Execute
        Else
            MsgBox "Data berhasil disimpan.. ", vbInformation, "Informasi"
            .Execute
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    cmdBatal_Click
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNama.SetFocus
End Sub

Private Sub dgTit_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgTit
    WheelHook.WheelHook dgTit
End Sub

Private Sub dgTit_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Errload
    If dgTit.ApproxCount = 0 Then Exit Sub
    txtkode.Text = dgTit.Columns(0).Value
    txtNama.Text = dgTit.Columns(1).Value
    chkStatus.Value = dgTit.Columns(4).Value
    txtKdExt.Text = dgTit.Columns(2).Value
    txtNamaExt.Text = dgTit.Columns(3).Value
    
    Exit Sub
Errload:
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo hell
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call cmdBatal_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Sub clear()
    On Error Resume Next
    txtkode.Text = ""
    txtNama.Text = ""
    txtNama.SetFocus
    txtKdExt.Text = ""
    chkStatus.Value = 1
    txtNamaExt.Text = ""
End Sub

Sub tampilData()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "select KdJenjang as Kode, NamaJenjangJabatan as [Jenjang Jabatan Fungsional], KodeExternal,NamaExternal,StatusEnabled from JenjangJabatan order by NamaJenjangJabatan "
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgTit.DataSource = rs
    dgTit.Columns(1).Width = 3500
    dgTit.Columns(4).Width = 1300
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtNamaExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
