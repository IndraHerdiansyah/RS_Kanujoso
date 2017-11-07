VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9b.ocx"
Begin VB.Form frmDataPersonalPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Personal Pegawai"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmDataPersonalPegawai.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10725
   Begin VB.TextBox txtJabatan 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1080
      TabIndex        =   32
      Top             =   7920
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   19
      Top             =   960
      Width           =   10695
      Begin VB.ComboBox cboStatusAktif 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmDataPersonalPegawai.frx":0CCA
         Left            =   7800
         List            =   "frmDataPersonalPegawai.frx":0CD4
         TabIndex        =   10
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtIdPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   240
         MaxLength       =   10
         TabIndex        =   20
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaLengkap 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   600
         Width           =   3015
      End
      Begin VB.ComboBox cboJK 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmDataPersonalPegawai.frx":0CDE
         Left            =   4800
         List            =   "frmDataPersonalPegawai.frx":0CE8
         TabIndex        =   1
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtTempatLahir 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6360
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtnip 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2040
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo dcPendidikan 
         Height          =   330
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcGolongan 
         Height          =   330
         Left            =   5520
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPangkat 
         Height          =   330
         Left            =   2760
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcJenisPegawai 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTTgl 
         Height          =   330
         Left            =   9000
         TabIndex        =   3
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93192193
         CurrentDate     =   38065
         MinDate         =   79
      End
      Begin MSDataListLib.DataCombo dcJabatan 
         Height          =   330
         Left            =   7320
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Status Aktif"
         Height          =   195
         Left            =   7800
         TabIndex        =   33
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID Pegawai"
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pegawai"
         Height          =   210
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nama Lengkap"
         Height          =   210
         Left            =   1680
         TabIndex        =   29
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   4800
         TabIndex        =   28
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir"
         Height          =   210
         Left            =   6360
         TabIndex        =   27
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Lahir"
         Height          =   210
         Left            =   9000
         TabIndex        =   26
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pangkat"
         Height          =   210
         Left            =   2760
         TabIndex        =   25
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Golongan"
         Height          =   210
         Left            =   5520
         TabIndex        =   24
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan"
         Height          =   210
         Left            =   7320
         TabIndex        =   23
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Pendidikan"
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "NIP"
         Height          =   210
         Left            =   4800
         TabIndex        =   21
         Top             =   1800
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   18
      Top             =   6960
      Width           =   10695
      Begin VB.CommandButton cmdAlamat 
         Caption         =   "&Alamat Pegawai"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1695
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
         Height          =   330
         Left            =   4920
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdUbah 
         Caption         =   "&Ubah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6360
         TabIndex        =   15
         Top             =   240
         Width           =   1335
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
         Height          =   330
         Left            =   7800
         TabIndex        =   16
         Top             =   240
         Width           =   1335
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
         Height          =   330
         Left            =   9240
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "&Baru"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid dgDataPegawai 
      Height          =   3135
      Left            =   0
      TabIndex        =   11
      Top             =   3720
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   34
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
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataPersonalPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8880
      Picture         =   "frmDataPersonalPegawai.frx":36C3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPersonalPegawai.frx":4BB1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataPersonalPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SetComboJenisPegawai()
    Set rs = Nothing
    rs.Open "Select * from JenisPegawai", dbConn, , adLockOptimistic
    Set dcJenisPegawai.RowSource = rs
    dcJenisPegawai.ListField = rs.Fields(2).Name
    dcJenisPegawai.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub setcomboPangkat()
    Set rs = Nothing
    rs.Open "Select * from Pangkat", dbConn, , adLockOptimistic
    Set dcPangkat.RowSource = rs
    dcPangkat.ListField = rs.Fields(1).Name
    dcPangkat.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboGolonganPegawai()
    Set rs = Nothing
    rs.Open "Select * from GolonganPegawai", dbConn, , adLockOptimistic
    Set dcGolongan.RowSource = rs
    dcGolongan.ListField = rs.Fields(1).Name
    dcGolongan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboJabatan()
    Set rs = Nothing
    rs.Open "select * from Jabatan", dbConn, , adLockOptimistic
    Set dcJabatan.RowSource = rs
    dcJabatan.ListField = rs.Fields(1).Name
    dcJabatan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Sub SetComboPendidikan()
    Set rs = Nothing
    rs.Open "select * from Pendidikan", dbConn, , adLockOptimistic
    Set dcPendidikan.RowSource = rs
    dcPendidikan.ListField = rs.Fields(1).Name
    dcPendidikan.BoundColumn = rs.Fields(0).Name
    Set rs = Nothing
End Sub

Private Sub cmdBatal_Click()
    Call kosong
    Call tombolsimpan
    txtNamaLengkap.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    Set rs = Nothing
    rs.Open "select * from DataPegawai where IdPegawai='" & txtidpegawai.Text & "' ", dbConn, adOpenStatic, adLockReadOnly
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        strSQL = "delete from DataPegawai where IdPegawai='" & txtidpegawai.Text & "'"
        dbConn.Execute strSQL
        MsgBox "Data Sukses Dihapus !", vbOKOnly, "Informasi"
        Call setdgDataPegawai
        Call kosong
        Call tombolsimpan
    End If
    Exit Sub
hell:
    MsgBox "Penghapusan Gagal, Data Sudah Terpakai !", vbOKOnly, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
    If dcJenisPegawai.Text = "" Then
        MsgBox "Jenis Pegawai harus diisi !", vbOKOnly, "Informasi"
        dcJenisPegawai.SetFocus
    ElseIf txtNamaLengkap.Text = "" Then
        MsgBox "Nama Lengkap harus diisi !", vbOKOnly, "Informasi"
        txtNamaLengkap.SetFocus
    ElseIf cboJK.Text = "" Then
        MsgBox "Jenis Kelamin harus diisi !", vbOKOnly, "Informasi"
        cboJK.SetFocus
    ElseIf txtTempatLahir.Text = "" Then
        MsgBox "Tempat Lahir harus diisi !", vbOKOnly, "Informasi"
        txtTempatLahir.SetFocus
    ElseIf dcPangkat.Text = "" Then
        MsgBox "Nama Pangkat harus diisi !", vbOKOnly, "Informasi"
        dcPangkat.SetFocus
    ElseIf dcGolongan.Text = "" Then
        MsgBox "Nama Golongan harus diisi !", vbOKOnly, "Informasi"
        dcGolongan.SetFocus
    ElseIf dcJabatan.Text = "" Then
        MsgBox "Jabatan harus diisi !", vbOKOnly, "Informasi"
        dcJabatan.SetFocus
    ElseIf dcPendidikan.Text = "" Then
        MsgBox "Pendidikan harus diisi !", vbOKOnly, "Informasi"
        dcPendidikan.SetFocus
    ElseIf txtNIP.Text = "" Then
        MsgBox "NIP harus diisi !", vbOKOnly, "Informasi"
        txtNIP.SetFocus
    Else
        Dim adoCommand As New ADODB.Command
        With adoCommand
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            If txtidpegawai.Text <> "" Then
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtidpegawai.Text)
            Else
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, Null)
            End If
            .Parameters.Append .CreateParameter("KdJenisPegawai", adChar, adParamInput, 3, Trim(dcJenisPegawai.BoundText))
            .Parameters.Append .CreateParameter("NamaLengkap", adVarChar, adParamInput, 50, txtNamaLengkap.Text)
            .Parameters.Append .CreateParameter("JenisKelamin", adChar, adParamInput, 1, Trim(cboJK.Text))
            .Parameters.Append .CreateParameter("TempatLahir", adVarChar, adParamInput, 50, txtTempatLahir.Text)
            .Parameters.Append .CreateParameter("TglLahir", adDate, adParamInput, , Format(DTTgl.Value, "yyyy/MM/dd"))
            .Parameters.Append .CreateParameter("KdPangkat", adVarChar, adParamInput, 2, dcPangkat.BoundText)
            .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, dcGolongan.BoundText)
            .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, dcJabatan.BoundText)
            .Parameters.Append .CreateParameter("KdPendidikanTerakhir", adChar, adParamInput, 2, dcPendidikan.BoundText)
            .Parameters.Append .CreateParameter("NIP", adVarChar, adParamInput, 10, txtNIP.Text)
            .Parameters.Append .CreateParameter("StatusAktif", adChar, adParamInput, 1, cboStatusAktif.Text)
            .Parameters.Append .CreateParameter("OutputIdPegawai", adChar, adParamOutput, 10, Null)
            
            .ActiveConnection = dbConn
            .CommandText = "Proc_D_GenerateIdPegawai"
            .CommandType = adCmdStoredProc
            .Execute
            
            If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                MsgBox "Ada kesalahan dalam penyimpanan Data Personal pegawai", vbCritical
                Call deleteADOCommandParameters(adoCommand)
                Set adoCommand = Nothing
                Exit Sub
            Else
                txtidpegawai.Text = .Parameters("OutputIdPegawai").Value
                MsgBox "Penyimpanan Data Personal Pegawai sukses", vbInformation
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With
    End If
    Call setdgDataPegawai
    Call kosong
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdubah_Click()
    If txtidpegawai.Text = "" Then
        MsgBox "Id Pegawai harus diisi !", vbOKOnly, "Informasi"
        txtidpegawai.SetFocus
    ElseIf dcJenisPegawai.Text = "" Then
        MsgBox "Jenis Pegawai harus diisi !", vbOKOnly, "Informasi"
        dcJenisPegawai.SetFocus
    ElseIf txtNamaLengkap.Text = "" Then
        MsgBox "Nama Lengkap harus diisi !", vbOKOnly, "Informasi"
        txtNamaLengkap.SetFocus
    ElseIf cboJK.Text = "" Then
        MsgBox "Jenis Kelamin harus diisi !", vbOKOnly, "Informasi"
        cboJK.SetFocus
    ElseIf txtTempatLahir.Text = "" Then
        MsgBox "Tempat Lahir harus diisi !", vbOKOnly, "Informasi"
        txtTempatLahir.SetFocus
    ElseIf dcPangkat.Text = "" Then
        MsgBox "Nama Pangkat harus diisi !", vbOKOnly, "Informasi"
        dcPangkat.SetFocus
    ElseIf dcGolongan.Text = "" Then
        MsgBox "Nama Golongan harus diisi !", vbOKOnly, "Informasi"
        dcGolongan.SetFocus
    ElseIf dcJabatan.Text = "" Then
        MsgBox "Jabatan harus diisi !", vbOKOnly, "Informasi"
        dcJabatan.SetFocus
    ElseIf dcPendidikan.Text = "" Then
        MsgBox "Pendidikan harus diisi !", vbOKOnly, "Informasi"
        dcPendidikan.SetFocus
    ElseIf txtNIP.Text = "" Then
        MsgBox "NIP harus diisi !", vbOKOnly, "Informasi"
        txtNIP.SetFocus
    Else
        cboJK.Text = Left(cboJK.Text, 1)
        Set rs = Nothing
        rs.Open "select * from DataPegawai where IdPegawai='" & txtidpegawai.Text & "' ", dbConn, adOpenStatic, adLockReadOnly
        If rs.RecordCount <> 0 Then
            Set rs = Nothing
            strSQL = "update DataPegawai set KdJenisPegawai='" & dcJenisPegawai.BoundText & "', NamaLengkap='" & txtNamaLengkap.Text & "', JenisKelamin='" & cboJK.Text & "', " _
                & "  TempatLahir='" & txtTempatLahir.Text & "', TglLahir= CONVERT(DateTime, '" & Format(DTTgl, "yyyy/mm/dd HH:mm:ss") & "', 102), KdPangkat='" & dcPangkat.BoundText & "', KdGolongan='" & dcGolongan.BoundText & "', kdJabatan='" & dcJabatan.BoundText & "'" _
                & "  , KdPendidikanTerakhir='" & dcPendidikan.BoundText & "', nip = '" & txtNIP.Text & "'  where IdPegawai='" & txtidpegawai.Text & "' "
            MsgBox "Data Sukses Diubah !", vbOKOnly, "Informasi"
            dbConn.Execute strSQL
            Call setdgDataPegawai
            Call kosong
            Call tombolsimpan
        Else
            MsgBox "Ubah Data Gagal", vbOKOnly, "Informasi"
            Exit Sub
        End If
    End If
End Sub

Private Sub dgDataPegawai_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call dgPegawai_KeyPress(13)
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call tombolubahapus
        txtidpegawai.Enabled = False
        txtidpegawai.Text = dgDataPegawai.Columns(0)
        dcJenisPegawai.Text = dgDataPegawai.Columns(1)
        txtNamaLengkap.Text = dgDataPegawai.Columns(2)
        cboJK.Text = dgDataPegawai.Columns(3)
        txtTempatLahir.Text = dgDataPegawai.Columns(4)
        If dgDataPegawai.Columns(5) = "" Then
            DTTgl.Value = Now
        Else
            DTTgl.Value = dgDataPegawai.Columns(5)
        End If
        dcPangkat.Text = dgDataPegawai.Columns(6)
        dcGolongan.Text = dgDataPegawai.Columns(7)
        dcJabatan.Text = dgDataPegawai.Columns(8)
        dcPendidikan.Text = dgDataPegawai.Columns(9)
        txtNIP.Text = dgDataPegawai.Columns(10)
    End If
End Sub

Private Sub DTTgl_Change()
    DTTgl.MaxDate = Now
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call SetComboJenisPegawai
    Call setcomboPangkat
    Call SetComboGolonganPegawai
    Call setdgDataPegawai
    Call SetComboJabatan
    Call SetComboPendidikan
    Call kosong
    Call tombolsimpan
End Sub

Sub setdgDataPegawai()
    Set rs = Nothing
    rs.Open "select * from V_JP_DataPegawai", dbConn, adOpenStatic, adLockReadOnly
    Set dgDataPegawai.DataSource = rs
    With dgDataPegawai
        .Columns(0).Width = 1400
        .Columns(1).Width = 1450
        .Columns(2).Width = 3000
        .Columns(3).Width = 1200
        .Columns(4).Width = 1200
        .Columns(5).Width = 1200
        .Columns(6).Width = 1700
        .Columns(7).Width = 1100
        .Columns(8).Width = 2500
        .Columns(9).Width = 1000
        .Columns(10).Width = 1000
        .Columns(0).Caption = "ID Pegawai"
        .Columns(1).Caption = "Jenis Pegawai"
        .Columns(2).Caption = "Nama Lengkap"
        .Columns(3).Caption = "Jns Kelamin"
        .Columns(4).Caption = "Tmpt Lahir"
        .Columns(5).Caption = "Tgl Lahir"
        .Columns(6).Caption = "Pangkat"
        .Columns(7).Caption = "Golongan"
        .Columns(8).Caption = "Jabatan"
        .Columns(9).Caption = "Pendidikan"
        .Columns(10).Caption = "NIP"
    End With
End Sub

Sub kosong()
    txtidpegawai.Text = ""
    dcJenisPegawai.Text = ""
    txtNamaLengkap.Text = ""
    cboJK.Text = ""
    txtTempatLahir.Text = ""
    DTTgl.Value = Date
    dcPangkat.Text = ""
    dcGolongan.Text = ""
    dcJabatan.Text = ""
    dcPendidikan.Text = ""
    txtNIP.Text = ""
    cboStatusAktif.Text = ""
End Sub

Sub tombolubahapus()
    cmdHapus.Enabled = True
    cmdubah.Enabled = True
    cmdSimpan.Enabled = False
End Sub

Sub tombolsimpan()
    cmdHapus.Enabled = False
    cmdubah.Enabled = False
    cmdSimpan.Enabled = True
End Sub

Private Sub txtIDPegawai_KeyPress(KeyAscii As Integer)
    Dim kelamin As String
    If KeyAscii = 13 Then
        Set rs = Nothing
        rs.Open "select * from V_JP_DataPegawai where IdPegawai='" & txtidpegawai.Text & "' ", dbConn, adOpenStatic, adLockReadOnly
        If rs.RecordCount <> 0 Then
            dcJenisPegawai.Text = rs.Fields("JenisPegawai").Value
            txtNamaLengkap.Text = rs.Fields("NamaLengkap").Value
            cboJK = rs.Fields("JenisKelamin").Value
            txtTempatLahir.Text = rs.Fields("TempatLahir").Value
            DTTgl.Value = rs.Fields("TglLahir").Value
            dcPangkat.Text = rs.Fields("NamaPangkat").Value
            dcGolongan.Text = rs.Fields("NamaGolongan").Value
            dcJabatan.Text = rs.Fields("Jabatan").Value
            dcPendidikan.Text = rs.Fields("Pendidikan").Value
            txtNIP.Text = rs.Fields("nip").Value
            txtidpegawai.Enabled = False
            Call tombolubahapus
        Else
            MsgBox "Data yang anda cari tidak ada dalam tabel.....!", vbInformation, "Pemberitahuan"
            Call kosong
            Call tombolsimpan
        End If
        Set rs = Nothing
    End If
End Sub

Private Sub txtNamaLengkap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cboJK.SetFocus
End Sub

Private Sub txtNamaLengkap_LostFocus()
    txtNamaLengkap.Text = StrConv(txtNamaLengkap.Text, vbProperCase)
End Sub

Private Sub cboJK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTempatLahir.SetFocus
End Sub

Private Sub DTTgl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJenisPegawai.SetFocus
End Sub

Private Sub dcJenisPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPangkat.SetFocus
End Sub

Private Sub dcPangkat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcGolongan.SetFocus
End Sub

Private Sub dcGolongan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcJabatan.SetFocus
End Sub

Private Sub dcJabatan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPendidikan.SetFocus
End Sub

Private Sub txtTempatLahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DTTgl.SetFocus
    End If
End Sub

Private Sub cboJK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTempatLahir.SetFocus
    End If
End Sub

Private Sub cboJK_LostFocus()
    cboJK.Text = Left(cboJK.Text, 1)
End Sub

Private Sub cmdAlamat_Click()
    frmDataAlamatPegawai.Show
End Sub

Private Sub dcGolongan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcJabatan.SetFocus
    End If
End Sub

Private Sub dcJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcPendidikan.SetFocus
    End If
End Sub

Private Sub dcJenisPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcPangkat.SetFocus
    End If
End Sub

Private Sub dcPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dcGolongan.SetFocus
    End If
End Sub

Private Sub dcPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNIP.SetFocus
    End If
End Sub

Private Sub txtNIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStatusAktif.SetFocus
    End If
End Sub

Private Sub cboStatusAktif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSimpan.SetFocus
    End If
End Sub




