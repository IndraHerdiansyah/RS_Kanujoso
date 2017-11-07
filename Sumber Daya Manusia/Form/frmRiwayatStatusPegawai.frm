VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatStatusPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Status Pegawai"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmRiwayatStatusPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   10695
   Begin VB.CommandButton cmdcetaksuratlangsung 
      Caption         =   "&Cetak Surat"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   8040
      TabIndex        =   11
      Top             =   7080
      Width           =   1215
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
      Left            =   6720
      TabIndex        =   10
      Top             =   7080
      Width           =   1215
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
      Left            =   9360
      TabIndex        =   12
      Top             =   7080
      Width           =   1215
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
      Left            =   5400
      TabIndex        =   9
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtkode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3480
      MaxLength       =   50
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   10455
      Begin VB.TextBox txtTTD 
         Appearance      =   0  'Flat
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
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1320
         Width           =   5535
      End
      Begin VB.TextBox txtNoSK 
         Appearance      =   0  'Flat
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
         Left            =   240
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
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
         Left            =   240
         MaxLength       =   200
         TabIndex        =   7
         Top             =   2760
         Width           =   9975
      End
      Begin VB.TextBox txtalasan 
         Appearance      =   0  'Flat
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
         Left            =   240
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2040
         Width           =   9975
      End
      Begin MSDataListLib.DataCombo dcStatus 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtnostatus 
         Height          =   375
         Left            =   23280
         TabIndex        =   20
         Top             =   5501
         Width           =   150
      End
      Begin MSComCtl2.DTPicker dtpTglAwal 
         Height          =   330
         Left            =   3000
         TabIndex        =   1
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   122224640
         UpDown          =   -1  'True
         CurrentDate     =   39282
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   330
         Left            =   5280
         TabIndex        =   2
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy "
         Format          =   122421248
         UpDown          =   -1  'True
         CurrentDate     =   39282
      End
      Begin MSComCtl2.DTPicker dtpTglSurat 
         Height          =   330
         Left            =   7800
         TabIndex        =   3
         Top             =   600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   582
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   122421248
         UpDown          =   -1  'True
         CurrentDate     =   39282
      End
      Begin VB.Label Label1 
         Caption         =   "Tgl. Akhir"
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
         Left            =   5280
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Tgl. SK"
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
         Left            =   7800
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanda Tangan SK"
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
         Left            =   4680
         TabIndex        =   24
         Top             =   1080
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. SK"
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
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Tgl. Mulai"
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
         Left            =   3000
         TabIndex        =   19
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Alasan Keperluan"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Status"
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
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   840
      End
   End
   Begin MSDataGridLib.DataGrid dgRiwayatStatus 
      Height          =   2295
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4048
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
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   18
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
   Begin VB.TextBox txtnoriwayat 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   120
      MaxLength       =   50
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   945
      Left            =   8640
      Picture         =   "frmRiwayatStatusPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2115
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatStatusPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatStatusPegawai.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatStatusPegawai.frx":470E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7680
      Picture         =   "frmRiwayatStatusPegawai.frx":70CF
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   1875
   End
End
Attribute VB_Name = "frmRiwayatStatusPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadGrid
    dcStatus.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If txtNoRiwayat.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatStatusPegawai WHERE NoRiwayat='" & txtNoRiwayat.Text & "' "
    dbConn.Execute strSQL
    If sp_Riwayat("D") = False Then Exit Sub

    MsgBox "Data Berhasil Dihapus ", vbInformation, "Informasi"
    Call cmdBatal_Click

    Dim tempNoUrutX As String
    Dim tempKdStatusX As String
    Dim tempTglAwal As Date
    strSQL = "select max(NoRiwayat) from RiwayatStatusPegawai Where IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempNoUrutX = rs(0).Value
    strSQL = "select KdStatus, TglAwal from RiwayatStatusPegawai where NoRiwayat='" & tempNoUrutX & "' and IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempKdStatusX = rs(0).Value
    tempTglAwal = rs(1).Value
    strSQL = "update DataCurrentPegawai set KdStatus='" & tempKdStatusX & "' where IdPegawai='" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    If tempKdStatusX = "01" Then
        strSQL = "update DataPegawai set TglKeluar=Null where IdPegawai='" & mstrIdPegawai & "'"
        Call msubRecFO(rs, strSQL)
    Else
        strSQL = "update DataPegawai set TglKeluar='" & Format(tempTglAwal, "yyyy/MM/dd") & "' where IdPegawai='" & mstrIdPegawai & "'"
        Call msubRecFO(rs, strSQL)
    End If
    Exit Sub
errHapus:
End Sub

Private Function sp_Riwayat(f_Status) As Boolean
    On Error GoTo hell
    sp_Riwayat = True
    Set dbcmd = New ADODB.Command
    With dbcmd

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If txtNoRiwayat = "" Then
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
        End If

        .Parameters.Append .CreateParameter("TglRiwayat", adDate, adParamInput, , Format(Now, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("OutputNoRiwayat", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Riwayat"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data riwayat ", vbCritical, "Validasi"
            sp_Riwayat = False
        End If
        txtNoRiwayat.Text = .Parameters("OutputNoRiwayat").Value
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatStatusPegawai() As Boolean
    On Error GoTo hell
    sp_RiwayatStatusPegawai = True
    Set adoComm = New ADODB.Command
    With adoComm

        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, txtNoRiwayat.Text)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatus.BoundText)
        .Parameters.Append .CreateParameter("TglAwal", adDate, adParamInput, , Format(dtpTglAwal.Value, "yyyy/MM/dd"))

        If IsNull(dtpTglAkhir.Value) Then
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(dtpTglAkhir.Value, "yyyy/MM/dd"))
        End If
        If IsNull(dtpTglSurat.Value) Then
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSurat.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, IIf(txtNoSK.Text = "", Null, Trim(txtNoSK.Text)))
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 50, IIf(txtTTD.Text = "", Null, Trim(txtTTD.Text)))
        .Parameters.Append .CreateParameter("AlasanKeperluan", adVarChar, adParamInput, 100, IIf(txtalasan.Text = "", Null, Trim(txtalasan.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)

        .ActiveConnection = dbConn
        .CommandText = "Add_RiwayatStatusPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data riwayat status pegawai", vbCritical, "Validasi"
            sp_RiwayatStatusPegawai = False
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Sub cmdSimpan_Click()
    On Error GoTo Errload
    If dcStatus.Text <> "" Then
        If Periksa("datacombo", dcStatus, "Status Tidak Terdaftar") = False Then Exit Sub
    End If
    
    If Periksa("datacombo", dcStatus, "Silahkan isi status ") = False Then Exit Sub
    If Periksa("text", txtalasan, "Silahkan isi alasan keperluan ") = False Then Exit Sub

    If sp_Riwayat("A") = False Then Exit Sub
    If sp_RiwayatStatusPegawai() = False Then Exit Sub

    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"

    Call cmdBatal_Click

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatStatus
    Unload Me
End Sub

Private Sub dcStatus_Change()
'    If dcStatus.BoundText = "07" Then
'        dtpTglAkhir.Value = Format(dtpTglAkhir.Value, "MM")
'
'        dtpTglAkhir.Enabled = False
'    Else
'        dtpTglAkhir.Enabled = True
'    End If
End Sub

Private Sub dcStatus_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then dtpTglAwal.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcStatus.Text)) = 0 Then dtpTglAwal.SetFocus: Exit Sub
        If dcStatus.MatchedWithList = True Then dtpTglAwal.SetFocus: Exit Sub
        strSQL = "select KdStatus,Status from StatusPegawai WHERE (Status LIKE '%" & dcStatus.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcStatus.BoundText = rs(0).Value
        dcStatus.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgRiwayatStatus_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgRiwayatStatus.ApproxCount = 0 Then Exit Sub
    dcStatus.BoundText = dgRiwayatStatus.Columns("KdStatus").Value
    dtpTglAwal.Value = dgRiwayatStatus.Columns("Tgl. Awal").Value
    If IsNull(dgRiwayatStatus.Columns("Tgl. Akhir").Value) Then dtpTglAkhir.Value = Null Else dtpTglAkhir.Value = dgRiwayatStatus.Columns("Tgl. Akhir").Value
    If IsNull(dgRiwayatStatus.Columns("TglSK").Value) Then dtpTglSurat.Value = "" Else dtpTglSurat.Value = dgRiwayatStatus.Columns("TglSK").Value
    txtalasan.Text = dgRiwayatStatus.Columns("Alasan Keperluan").Value
    If IsNull(dgRiwayatStatus.Columns("Keterangan").Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgRiwayatStatus.Columns("Keterangan").Value
    If IsNull(dgRiwayatStatus.Columns("NoSK").Value) Then txtNoSK.Text = "" Else txtNoSK.Text = dgRiwayatStatus.Columns("NoSK").Value
    If IsNull(dgRiwayatStatus.Columns("TandaTanganSK").Value) Then txtTTD.Text = "" Else txtTTD.Text = dgRiwayatStatus.Columns("TandaTanganSK").Value
    txtNoRiwayat.Text = dgRiwayatStatus.Columns("NoRiwayat").Value
End Sub

Private Sub dtpTglAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglAkhir.SetFocus
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTglSurat.SetFocus
End Sub

Private Sub dtpTglSurat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtNoSK.SetFocus
End Sub

Private Sub Form_Load()
    On Error GoTo Errload

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadDcSource
    Call subLoadGrid
    dtpTglAwal.Value = Now
    dtpTglAkhir.Value = Now
    dtpTglSurat.Value = Now

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub subLoadGrid()
    On Error GoTo hell
    Set rs = Nothing
    strSQL = "select * from V_RiwayatStatusPegawai_New WHERE ID='" & mstrIdPegawai & "' ORDER BY NoRiwayat"
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    Set dgRiwayatStatus.DataSource = rs
    With dgRiwayatStatus
        .Columns("NoRiwayat").Width = 0
        .Columns("KdStatus").Width = 0
    End With
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    dcStatus.Text = ""
    dtpTglAwal.Value = Now
    dtpTglAkhir.Value = Now
    txtalasan.Text = ""
    txtKeterangan.Text = ""
    txtNoRiwayat.Text = ""
    dtpTglSurat.Value = Now
    txtNoSK.Text = ""
    txtTTD.Text = ""
    txtNoRiwayat.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatStatus
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtAlasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Private Sub subLoadDcSource()
    strSQL = "select KdStatus,Status from StatusPegawai where StatusEnabled='1' order by Status"
    Call msubDcSource(dcStatus, rs, strSQL)
End Sub

Private Sub txtNoSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTTD.SetFocus
End Sub

Private Sub txtTTD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtalasan.SetFocus
End Sub
