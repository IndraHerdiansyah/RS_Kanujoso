VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTukarShift 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Tukar Shift"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9420
   Icon            =   "frmTukarShift.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9420
   Begin MSDataGridLib.DataGrid dgNamaPegawai2 
      Height          =   2055
      Left            =   4800
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin MSDataGridLib.DataGrid dgNamaPegawai 
      Height          =   2055
      Left            =   -120
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
   Begin MSDataListLib.DataCombo dcShift 
      Height          =   315
      Left            =   2040
      TabIndex        =   18
      Top             =   3480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   1680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Format          =   57868288
      CurrentDate     =   39332
   End
   Begin VB.TextBox txtPegPengganti 
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   13
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtIDPegawai 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   8
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtNamaPegawai 
      BackColor       =   &H80000014&
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txttempatbertugas 
      BackColor       =   &H80000014&
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox txtjabatan 
      BackColor       =   &H80000014&
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
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Update"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   4440
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
      Left            =   3600
      TabIndex        =   2
      Top             =   4440
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
      Left            =   2400
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   3
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Shift Pengganti"
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
      Left            =   480
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Pegawai Pengganti"
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
      TabIndex        =   15
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. ID Pegawai"
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
      Left            =   480
      TabIndex        =   12
      Top             =   2400
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Lengkap"
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
      Index           =   2
      Left            =   480
      TabIndex        =   11
      Top             =   2040
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ruangan"
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
      Index           =   4
      Left            =   480
      TabIndex        =   10
      Top             =   3120
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jabatan"
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
      Index           =   8
      Left            =   480
      TabIndex        =   9
      Top             =   2760
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   945
      Left            =   7320
      Picture         =   "frmTukarShift.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2115
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmTukarShift.frx":21B8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmTukarShift.frx":3816
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7680
      Picture         =   "frmTukarShift.frx":61D7
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   1875
   End
End
Attribute VB_Name = "frmTukarShift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    dcStatus.SetFocus
End Sub

Private Function sp_Riwayat(f_Status) As Boolean
    sp_Riwayat = True
    Set adoComm = New ADODB.Command
    With adoComm
    
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If txtnoriwayat = "" Then
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        End If
        
        .Parameters.Append .CreateParameter("TglRiwayat", adDate, adParamInput, , Format(Now, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, txtkode.Text)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("OutputNoRiwayat", adChar, adParamOutput, 10, Null)
                
                        
        .ActiveConnection = dbConn
        .CommandText = "AUD_Riwayat"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Pegawai", vbCritical, "Validasi"
            sp_Riwayat = False
        'Else
            'If Not IsNull(.Parameters("Status").Value) Then Txt_IdPeg = .Parameters("Status").Value
            'mstrIdPegawai = Txt_IdPeg.Text
            'MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
        End If
        txtnoriwayat.Text = .Parameters("OutputNoRiwayat").Value
        'If Not IsNull(.Parameters("IdPegawai").Value) Then
        ' Update
            'MsgBox "Update data berhasil", vbInformation, "Informasi"
            '.Execute
        'Else
        ' Input data baru
            'MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
            '.Execute
        'End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function
Private Function sp_RiwayatStatusPegawai() As Boolean
    sp_RiwayatStatusPegawai = True
    Set adoComm = New ADODB.Command
    With adoComm
    
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtnoriwayat.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIDPegawai.Text)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatus.BoundText)
        .Parameters.Append .CreateParameter("TglAwal", adDate, adParamInput, , Format(dtpTglAwal.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(dtpTglAkhir.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("AlasanKeperluan", adVarChar, adParamInput, 100, txtalasan.Text)
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, txtKeterangan.Text))
       ' .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "U")
       ' .Parameters.Append .CreateParameter("OutPutNoRiwayat", adChar, adParamOutput, 2, Null)
                        
        .ActiveConnection = dbConn
        .CommandText = "Add_RiwayatStatusPegawai"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data Pegawai", vbCritical, "Validasi"
            sp_RiwayatStatusPegawai = False
        'Else
            'If Not IsNull(.Parameters("Status").Value) Then Txt_IdPeg = .Parameters("Status").Value
            'mstrIdPegawai = Txt_IdPeg.Text
            'MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
        End If
        
        'If Not IsNull(.Parameters("IdPegawai").Value) Then
        ' Update
            'MsgBox "Update data berhasil", vbInformation, "Informasi"
            '.Execute
        'Else
        ' Input data baru
            'MsgBox "Penyimpanan data berhasil", vbInformation, "Informasi"
            '.Execute
        'End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function

Private Sub cmdSimpan_Click()
On Error GoTo errload
    If Periksa("datacombo", dcStatus, "Status pegawai diisi!") = False Then Exit Sub
    If Periksa("text", txtalasan, "Alasan keperluan diisi!") = False Then Exit Sub
       
    If sp_Riwayat("U") = False Then Exit Sub
    If sp_RiwayatStatusPegawai() = False Then Exit Sub
    
    MsgBox "Data telah disimpan..", vbInformation, "Informasi"
        
    Call subClearData
    Call subLoadGrid

Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgRiwayatStatus_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgRiwayatStatus.ApproxCount = 0 Then Exit Sub
    
    dcStatus.Text = dgRiwayatStatus.Columns("Status").Value
    dtpTglAwal.Value = dgRiwayatStatus.Columns("Tgl. Awal").Value
    dtpTglAkhir.Value = dgRiwayatStatus.Columns("Tgl. Akhir").Value
    txtalasan.Text = dgRiwayatStatus.Columns("Alasan Keperluan").Value
    txtKeterangan.Text = dgRiwayatStatus.Columns("Keterangan").Value
    txtnoriwayat.Text = dgRiwayatStatus.Columns("NoRiwayat").Value
End Sub

Private Sub dtpTglAwal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub
Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub dgNamaPegawai_Click()
With dgNamaPegawai
        If .ApproxCount = 0 Then Exit Sub
        txtNamaPegawai.Text = dgNamaPegawai.Columns(0).Value
        .Visible = False
    End With
    'txtIDPegawai.SetFocus
End Sub

Private Sub dgNamaPegawai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call dgNamaPegawai_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
       
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    'Call subClearData
    Call subLoadDcSource
    'Call subLoadGrid
    dtpTglAwal.Value = Now
    dtpTglAkhir.Value = Now
    Exit Sub
errload:
Call msubPesanError
    'dtpTglAwal.Enabled = False
End Sub

Private Sub subClearData()
    dcStatus.Text = ""
    dtpTglAwal.Value = Now
    dtpTglAkhir.Value = Now
    txtalasan.Text = ""
    txtKeterangan.Text = ""
    txtnoriwayat.Text = ""
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtstatus_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then If dgNamaStatus.Visible = True Then dgNamaStatus.SetFocus
End Sub

Private Sub txtNamastatus_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then If dgNamaStatus.Visible = True Then dgNamaStatus.SetFocus Else txtalasan.SetFocus

End Sub

Private Sub subLoadDcSource()
On Error GoTo errload
    strSQL = "select KdShift,NamaShift from ShiftKerja order by NamaShift"
    Call msubDcSource(dcShift, rs, strSQL)
Exit Sub
errload:
    Call msubPesanError
End Sub

Private Sub txtNamaPegawai_Change()
    Call SubLoadGridNamaPegawai
End Sub

Private Sub SubLoadGridNamaPegawai()
On Error Resume Next
strSQL = "SELECT NamaLengkap, IdPegawai, NamaRuangan, NamaJabatan FROM v_DP_TukarShift where NamaLengkap LIKE '%" & txtNamaPegawai.Text & "%'"
Call msubRecFO(rs, strSQL)
Set dgNamaPegawai.DataSource = rs
With dgNamaPegawai
    .Columns("NamaLengkap").Width = 2500
    .Columns("NamaLengkap").Caption = "Nama"
    .Columns("IdPegawai").Width = 0
    .Columns("NamaRuangan").Width = 0
    .Columns("NamaJabatan").Width = 0
    txtIDPegawai.Text = .Columns("IdPegawai").Value
    If .Columns("NamaRuangan").Value = "" Then
        txttempatbertugas.Text = ""
    Else
        txttempatbertugas.Text = .Columns("NamaRuangan").Value
    End If
    If .Columns("NamaJabatan").Value = "" Then
        txtjabatan.Text = ""
    Else
        txtjabatan.Text = .Columns("NamaJabatan").Value
    End If
    .Visible = True
End With
Exit Sub
'errload:
'    Call msubPesanError
End Sub

Private Sub txtNamaPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then If dgNamaPegawai.Visible = True Then dgNamaPegawai.SetFocus
End Sub

Private Sub txtNamaPegawai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then If dgNamaPegawai.Visible = True Then dgNamaPegawai.SetFocus 'Else txtIDPegawai.SetFocus
End Sub

Private Sub txtPegPengganti_Change()
On Error Resume Next
strSQL = "SELECT NamaLengkap, IdPegawai, NamaRuangan, NamaJabatan FROM v_DP_TukarShift where NamaLengkap LIKE '%" & txtNamaPegawai.Text & "%'"
Call msubRecFO(rs, strSQL)
Set dgNamaPegawai2.DataSource = rs
With dgNamaPegawai2
    .Columns("NamaLengkap").Width = 2500
    .Columns("NamaLengkap").Caption = "Nama"
    .Columns("IdPegawai").Width = 0
    .Columns("NamaRuangan").Width = 2000
    .Columns("NamaRuangan").Caption = "Ruangan"
    .Columns("NamaJabatan").Width = 2500
    .Columns("NamaJabatan").Caption = "Jabatan"
    .Visible = True
End With
End Sub

Private Sub txtPegPengganti_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then If dgNamaPegawai2.Visible = True Then dgNamaPegawai2.SetFocus
End Sub

Private Sub txtPegPengganti_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then If dgNamaPegawai2.Visible = True Then dgNamaPegawai2.SetFocus 'Else txtIDPegawai.SetFocus
End Sub

Private Sub dgNamaPegawai2_Click()
With dgNamaPegawai2
        If .ApproxCount = 0 Then Exit Sub
        txtPegPengganti.Text = dgNamaPegawai2.Columns(0).Value
        .Visible = False
    End With
    'txtIDPegawai.SetFocus
End Sub

Private Sub dgNamaPegawai2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call dgNamaPegawai2_Click
End Sub
