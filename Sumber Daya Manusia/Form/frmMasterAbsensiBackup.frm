VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMasterAbsensiBackUp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Hari Libur"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmMasterAbsensiBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   6495
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   7560
      Width           =   1095
   End
   Begin VB.TextBox txtKdExtKlmpk 
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox chkStatusKlmpk 
      Alignment       =   1  'Right Justify
      Caption         =   "Status Enabled"
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtNamaExtKlmpk 
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
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox txtnamaharilibur 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtkdharilibur 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   7560
      Width           =   1095
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   12
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
   Begin MSDataGridLib.DataGrid dgHarilibur 
      Height          =   4695
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpTglLibur 
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   82640897
      CurrentDate     =   40100
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Kode External"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Nama External"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nama Hari Libur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   15
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Tanggal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4920
      TabIndex        =   14
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Kode"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   5040
      Picture         =   "frmMasterAbsensiBackup.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterAbsensiBackup.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterAbsensiBackup.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmMasterAbsensiBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sub subDcSource()
''On Error Resume Next
''   strSQL = "SELECT * FROM JenisHari order by JenisHari"
''   Call msubDcSource(dcJnsharilibur, rs, strSQL)
'End Sub

Sub sp_simpan()
'    Select Case sstDataPenunjang.Tab
'        Case 0 'jenis hari
'        Set dbcmd = New ADODB.Command
'           With dbcmd
'               .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
'               .Parameters.Append .CreateParameter("KdJenisHari", adChar, adParamInput, 2, txtkdjenishari.Text)
'               .Parameters.Append .CreateParameter("JenisHari", adVarChar, adParamInput, 20, txtnamajenishari.Text)
'               .Parameters.Append .CreateParameter("OutputKdJenisHari", adChar, adParamOutput, 2, Null)
'
'                .ActiveConnection = dbConn
'                .CommandText = "AU_JenisHari"
'                .CommandType = adCmdStoredProc
'                .Execute
'
'                If Not (.Parameters("return_value").Value = 0) Then
'                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
'                Else
'                   If Not IsNull(.Parameters("OutputKdJenisHari").Value) Then txtkdjenishari = .Parameters("OutputKdjenishari").Value
'               End If
'               Call deleteADOCommandParameters(dbcmd)
'            End With
'            cmdBatal_Click
'
'        Case 1 'Hari Libur
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdHariLibur", adVarChar, adParamInput, 3, txtkdharilibur.Text)
                .Parameters.Append .CreateParameter("TglHariLibur", adDate, adParamInput, , Format(dtpTglLibur.Value, "yyyy/MM/dd"))
                .Parameters.Append .CreateParameter("NamaHariLibur", adVarChar, adParamInput, 50, txtnamaharilibur.Text)
                .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 100, Null)
                .Parameters.Append .CreateParameter("KdJenisHari", adChar, adParamInput, 2, "01")
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExtKlmpk.Text)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNamaExtKlmpk.Text)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStatusKlmpk.Value)
                .Parameters.Append .CreateParameter("OutputKdharilibur", adVarChar, adParamOutput, 3, Null)
                    
                .ActiveConnection = dbConn
                .CommandText = "AU_HariLibur"
                .CommandType = adCmdStoredProc
                .Execute
              
                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                Else
                   If Not IsNull(.Parameters("OutputKdHariLibur").Value) Then txtkdharilibur = .Parameters("OutputKdharilibur").Value
               End If
               Call deleteADOCommandParameters(dbcmd)
            End With
            cmdBatal_Click
        
'    End Select
End Sub

Private Sub chkStatusKlmpk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExtKlmpk.SetFocus
End Sub

Private Sub cmdBatal_Click()
'    Select Case sstDataPenunjang.Tab
'        Case 0 'jenishari
            Call subKosong
            Call subLoadGridSource
'            cmdHapus.Enabled = True
'            cmdSimpan.Enabled = True
'        Case 1 'Hari Libur
'            Call subKosong
'            cmdHapus.Enabled = True
'            cmdSimpan.Enabled = True
'    End Select
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell:
'    On Error Resume Next
    If dgHarilibur.ApproxCount = 0 Then Exit Sub
    vLaporan = ""
    If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
    frmHariLibur.Show
hell:
End Sub

Private Sub cmdHapus_Click()
'    Select Case sstDataPenunjang.Tab
'        Case 0 'jenis hari
'            Set rs = Nothing
'
'            strSQL = "delete JenisHari where KdJenisHari = '" & txtkdjenishari & "'"
'            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
'            Set rs = Nothing
'
'        Case 1 'Hari Libur
           If Periksa("text", txtnamaharilibur, "Nama libur kosong") = False Then Exit Sub
           If MsgBox("Yakin data ini akan dihapus", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            Set rs = Nothing
            strSQL = "delete HariLibur where KdHariLibur = '" & txtkdharilibur.Text & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
            MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
'    End Select
    Call subLoadGridSource
    Call subKosong
End Sub

Private Sub cmdSimpan_Click()
'On Error GoTo errSimpan
'    Select Case sstDataPenunjang.Tab
'        Case 0 'jenis hari
'            If Periksa("text", txtnamajenishari, "Jenis Hari Harus diisi!!") = False Then Exit Sub
'            Call sp_simpan
'        Case 1  ' Hari Libur
            If Periksa("text", txtnamaharilibur, "Nama libur kosong") = False Then Exit Sub
            Call sp_simpan
            MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
'    End Select
    Call subLoadGridSource
Exit Sub
errSimpan:
    msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgHarilibur_Click()
WheelHook.WheelUnHook
        Set MyProperty = dgHarilibur
        WheelHook.WheelHook dgHarilibur
End Sub

'Private Sub dcJnsharilibur_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dtpTglLibur.SetFocus
'End Sub
'
'Private Sub dgjenishari_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    txtkdjenishari.Text = dgJenisHari.Columns(0).Value
'    txtnamajenishari.Text = dgJenisHari.Columns(1).Value
'End Sub

Private Sub dgHariLibur_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    txtkdharilibur.Text = dgHarilibur.Columns("kdharilibur").Value
    dtpTglLibur.Value = dgHarilibur.Columns("TglHariLibur").Value
    txtnamaharilibur.Text = dgHarilibur.Columns("NamaHariLibur").Value
    txtKdExtKlmpk.Text = dgHarilibur.Columns("KodeExternal").Value
    txtNamaExtKlmpk.Text = dgHarilibur.Columns("NamaExternal").Value
    chkStatusKlmpk.Value = dgHarilibur.Columns("StatusEnabled").Value
'   txtketerangan.Text = dgHarilibur.Columns(3).Value
'   dcJnsharilibur.Text = dgHarilibur.Columns(4).Value
End Sub

Private Sub dtpTglLibur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtKdExtKlmpk.SetFocus
End Sub



Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
'    Call subDcSource
'    sstDataPenunjang.Tab = 0
    Call subLoadGridSource
   
End Sub

Sub subKosong()
On Error Resume Next
'    Select Case sstDataPenunjang.Tab
'        Case 0 'jenishari
'            txtkdjenishari.Text = ""
'            txtnamajenishari.Text = ""
'        Case 1 'HariLibur
            txtkdharilibur.Text = ""
            'dcJnsharilibur.Text = ""
            dtpTglLibur.Value = Now
            txtnamaharilibur.Text = ""
            txtKdExtKlmpk.Text = ""
            txtNamaExtKlmpk.Text = ""
            chkStatusKlmpk.Value = 1
            'txtketerangan.Text = ""
'    End Select
End Sub

'Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
'    Call subDcSource
'    Call subLoadGridSource
'    Call cmdBatal_Click
'    dtpTglLibur.Value = Now
'End Sub

Sub subLoadGridSource()
'    On Error Resume Next
'    Select Case sstDataPenunjang.Tab
'           Case 0 ' jenishari
'            Set rs = Nothing
'            strSQL = "select * from jenishari"
'            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
'            Set dgJenisHari.DataSource = rs
'                dgJenisHari.Columns(0).DataField = rs(0).Name
'                dgJenisHari.Columns(1).DataField = rs(1).Name
'                dgJenisHari.Columns(0).Caption = "Kode"
'                dgJenisHari.Columns(1).Caption = "Jenis Hari"
'                dgJenisHari.Columns(0).Width = 1000
'                dgJenisHari.Columns(1).Width = 3200
'            Set rs = Nothing
'
'          Case 1  'Hari Libur
          Set rs = Nothing
            strSQL = "SELECT * " & _
                    " FROM HariLibur "
            Call msubRecFO(rs, strSQL)
            Set dgHarilibur.DataSource = rs
                dgHarilibur.Columns(0).Width = 1000
                dgHarilibur.Columns(1).Width = 1000
                dgHarilibur.Columns(2).Width = 1000
                dgHarilibur.Columns(3).Width = 0
                dgHarilibur.Columns(4).Width = 0
                dgHarilibur.Columns(7).Width = 1000
                dtpTglLibur.DataField = Now
'            Set rs = Nothing
'    End Select
End Sub

Private Sub txtKdExtKlmpk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkStatusKlmpk.SetFocus
End Sub

Private Sub txtNamaExtKlmpk_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

'Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then cmdSimpan.SetFocus
'End Sub

Private Sub txtnamaharilibur_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglLibur.SetFocus
End Sub

'Private Sub txtnamajenishari_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then Me.cmdSimpan.SetFocus
'End Sub

