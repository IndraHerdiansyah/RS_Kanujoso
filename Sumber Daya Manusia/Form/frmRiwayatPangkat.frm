VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRiwayatPangkat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Pangkat Pegawai"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "frmRiwayatPangkat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10455
   Begin VB.CommandButton Command3 
      Caption         =   "Delete File"
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
      Left            =   1440
      TabIndex        =   28
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ListBox lstFile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   10215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open File"
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
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dgNamaPangkat 
      Height          =   2535
      Left            =   1320
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
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
      Left            =   4680
      TabIndex        =   7
      Top             =   6360
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
      Height          =   375
      Left            =   9000
      TabIndex        =   10
      Top             =   6360
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
      Height          =   375
      Left            =   6120
      TabIndex        =   8
      Top             =   6360
      Width           =   1335
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
      Left            =   7560
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame fraPerjalanan 
      Caption         =   "Riwayat Pangkat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   10215
      Begin VB.CommandButton Command2 
         Caption         =   "File Upload"
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
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtGol 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtKdPangkat 
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtNoUrut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtNoSK 
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   4920
         MaxLength       =   30
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtNamaPangkat 
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtTandaTanganSK 
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1320
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   7800
         TabIndex        =   3
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
         CheckBox        =   -1  'True
         Format          =   81199104
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin VB.TextBox txtKeterangan 
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
         ForeColor       =   &H80000006&
         Height          =   330
         Left            =   5280
         MaxLength       =   200
         TabIndex        =   6
         Top             =   1320
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker dtpTMT 
         Height          =   330
         Left            =   240
         TabIndex        =   4
         Top             =   1320
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
         CheckBox        =   -1  'True
         Format          =   120127488
         UpDown          =   -1  'True
         CurrentDate     =   36872
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   7800
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gol."
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
         Left            =   4200
         TabIndex        =   23
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Berlaku"
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
         TabIndex        =   21
         Top             =   1080
         Width           =   840
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
         Left            =   5280
         TabIndex        =   19
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. Urut"
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
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   4920
         TabIndex        =   15
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pangkat"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label12 
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
         Left            =   2520
         TabIndex        =   13
         Top             =   1080
         Width           =   1260
      End
   End
   Begin MSDataGridLib.DataGrid dgRiwayatPangkat 
      Height          =   2655
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4683
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
      TabIndex        =   17
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   8640
      Picture         =   "frmRiwayatPangkat.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmRiwayatPangkat.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRiwayatPangkat.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmRiwayatPangkat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub subLoadGridNamaPangkat()
    On Error GoTo errLoad
    strSQL = "SELECT dbo.Pangkat.KdPangkat, dbo.Pangkat.NamaPangkat AS Pangkat, dbo.GolonganPegawai.NamaGolongan AS Gol" & _
    " FROM dbo.Pangkat INNER JOIN " & _
    "dbo.GolonganPegawai ON dbo.Pangkat.KdGolongan = dbo.GolonganPegawai.KdGolongan where dbo.Pangkat.NamaPangkat LIKE '%" & txtNamaPangkat.Text & "%' order by dbo.Pangkat.NamaPangkat"
    Call msubRecFO(rs, strSQL)
    Set dgNamaPangkat.DataSource = rs
    With dgNamaPangkat
        .Columns("KdPangkat").Width = 0
        .Columns(1).Width = 2000
        .Columns(2).Width = 1000

        .Visible = True
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
    Call subClearData
    Call subLoadRiwayatPangkat
    txtNamaPangkat.SetFocus
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    Dim tempNoUrutX As String
    Dim tempKdPangkatX As String
    If txtNoUrut.Text = "" Then Exit Sub
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM RiwayatPangkat WHERE IdPegawai='" & mstrIdPegawai & "' AND NoUrut='" & txtNoUrut.Text & "'"
    dbConn.Execute strSQL
    strSQL = "select max(NoUrut) from V_BantuUpdatePangkatToCurrent Where IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempNoUrutX = rs(0).Value
    strSQL = "select KdPangkat from Pangkat where NoUrut='" & tempNoUrutX & "'"
    Call msubRecFO(rs, strSQL)
    tempKdPangkatX = rs(0).Value
    strSQL = "update DataCurrentPegawai set KdPangkat='" & tempKdPangkatX & "' where IdPegawai='" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"

    Call cmdBatal_Click
    Exit Sub
errHapus:
    MsgBox "Data tidak dapat dihapus ", vbCritical, "Validasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo hell
    If Periksa("text", txtNamaPangkat, "Silahkan isi nama pangkat ") = False Then Exit Sub
    If dgNamaPangkat.Visible = True Then
       MsgBox "Nama Pangkat yang di inputkan Tidak ada di daftar.", vbCritical, "Validasi"
       txtNamaPangkat = "": txtNamaPangkat.SetFocus
       Exit Sub
    End If
    

'    If Periksa("text", txtGol, "Silahkan isi Golongan ") = False Then Exit Sub
    
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, Null)
        End If
        .Parameters.Append .CreateParameter("KdPangkat", adVarChar, adParamInput, 2, Trim(txtKdPangkat.Text))
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, IIf(txtNoSK.Text = "", Null, Trim(txtNoSK.Text)))
        If IsNull(dtpTglSK.Value) Then
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 30, IIf(txtTandaTanganSK.Text = "", Null, Trim(txtTandaTanganSK.Text)))
        .Parameters.Append .CreateParameter("Keterangan", adVarChar, adParamInput, 200, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawai)
        If IsNull(dtpTMT.Value) Then
            .Parameters.Append .CreateParameter("TMT", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TMT", adDate, adParamInput, , Format(dtpTMT.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 2, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RPangkat"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Pangkat", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            Exit Sub
        Else
            txtNoUrut.Text = .Parameters("OutputNoUrut").Value
            MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    
    
    'Upload file ____________________________________________________________________________________________________________________________
    Dim fso As New FileSystemObject
    Dim pathfile As String
    
    strSQL = "delete PathFileUpload where Jenis='Pangkat' and Kode='" & txtKdPangkat.Text & "'"
    Call msubRecFO(rs, strSQL)
    
    Dim anjing As Integer
    For anjing = 0 To lstFile.ListCount - 1
        pathfile = Replace(mstrPathFileSDM & "\" & "Pangkat_" & mstrIdPegawai & "_" & txtKdPangkat.Text & "\" & fso.GetFileName(lstFile.List(anjing)), "\", "/")
        strSQL = "insert into PathFileUpload values ('Pangkat','" & txtKdPangkat.Text & "','" & pathfile & "','')"
        Call msubRecFO(rs, strSQL)
        
        If fso.FolderExists(mstrPathFileSDM & "\" & "Pangkat_" & mstrIdPegawai & "_" & txtKdPangkat.Text) = False Then fso.CreateFolder mstrPathFileSDM & "\" & "Pangkat_" & mstrIdPegawai & "_" & txtKdPangkat.Text
        fso.CopyFile lstFile.List(anjing), mstrPathFileSDM & "\" & "Pangkat_" & mstrIdPegawai & "_" & txtKdPangkat.Text & "\" & fso.GetFileName(lstFile.List(anjing)), True
        lstFile.List(anjing) = mstrPathFileSDM & "\" & "Pangkat_" & mstrIdPegawai & "_" & txtKdPangkat.Text & "\" & fso.GetFileName(lstFile.List(anjing))
        If fso.FileExists(lstFile.List(anjing)) = False Then MsgBox "Upload file gagal..!", vbInformation, "..:.": Exit For
    Next
    Call Command2_Click
    '_________________________________________________________________________________________________________________________________________
    
    dgNamaPangkat.Visible = False
    Call subLoadRiwayatPangkat
    Call subClearData
    txtNamaPangkat.SetFocus
    Dim tempNoUrutX As String
    Dim tempKdPangkatX As String

    strSQL = "select max(NoUrut) from V_BantuUpdatePangkatToCurrent Where IdPegawai = '" & mstrIdPegawai & "'"
    Call msubRecFO(rs, strSQL)
    tempNoUrutX = rs(0).Value
    strSQL = "select KdPangkat from Pangkat where NoUrut='" & tempNoUrutX & "'"
    Call msubRecFO(rs, strSQL)
    
    
'    tempKdPangkatX = rs(0).Value
'    strSQL = "update DataCurrentPegawai set KdPangkat='" & tempKdPangkatX & "' where IdPegawai='" & mstrIdPegawai & "'"
'    Call msubRecFO(rs, strSQL)
    Exit Sub
hell:
    'Resume 0
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Call frmRiwayatPegawai.subLoadRiwayatPangkat
    Unload Me
End Sub

Private Sub Command1_Click()
'    Dim hwnd
'   Dim StartDoc
'   hwnd = apiFindWindow("OPUSAPP", "0")
'
'   StartDoc = ShellExecute(hwnd, "open", txtpathfile.Text, "", "C:\", 1)
On Error Resume Next
    Shell "cmd /c """ & lstFile.List(lstFile.ListIndex) & """, vbNormalFocus"
End Sub

Private Sub Command2_Click()
On Error Resume Next
    Command1.Visible = Not Command1.Visible
    lstFile.Visible = Not lstFile.Visible
    Command3.Visible = Not Command3.Visible
    
'    lstFile.clear
'
'Dim fl() As String
'Dim anjing As Integer
'Dim fso As New FileSystemObject
'
'    fl = Split(txtpathfile.Text, "|")
'    For anjing = 1 To 10
'        If fso.FileExists(fl(anjing)) = True Then
'            lstFile.AddItem fl(anjing)
'        End If
'    Next
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    
Dim fso As New FileSystemObject

    If InStr(1, lstFile.List(lstFile.ListIndex), mstrPathFileSDM) > 0 Then
        fso.DeleteFile lstFile.List(lstFile.ListIndex), True
        strSQL = "delete from PathFileUpload where Jenis='Pangkat' and Kode='" & txtKdPangkat.Text & "' and PathFile='" & Replace(lstFile.List(lstFile.ListIndex), "\", "/") & "'"
        Call msubRecFO(rs, strSQL)
    End If
    
    lstFile.RemoveItem lstFile.ListIndex
End Sub

Private Sub dgNamaPangkat_DblClick()
    On Error Resume Next
    With dgNamaPangkat
        If .ApproxCount = 0 Then Exit Sub
        txtKdPangkat.Text = dgNamaPangkat.Columns(0).Value
        txtNamaPangkat.Text = dgNamaPangkat.Columns(1).Value
        txtGol.Text = dgNamaPangkat.Columns(2).Value
        .Visible = False
    End With
    txtNoSK.SetFocus
End Sub

Private Sub dgNamaPangkat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgNamaPangkat_DblClick
End Sub

Private Sub dgRiwayatPangkat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgRiwayatPangkat.ApproxCount = 0 Then Exit Sub
    txtNoUrut.Text = dgRiwayatPangkat.Columns(1).Value
    txtKdPangkat.Text = dgRiwayatPangkat.Columns(2).Value
    txtNamaPangkat.Text = dgRiwayatPangkat.Columns(3).Value
    txtGol.Text = dgRiwayatPangkat.Columns(4).Value
    If IsNull(dgRiwayatPangkat.Columns(5).Value) Then txtNoSK.Text = "" Else txtNoSK.Text = dgRiwayatPangkat.Columns(5).Value
    If IsNull(dgRiwayatPangkat.Columns(6).Value) Then dtpTglSK.Value = Null Else dtpTglSK.Value = dgRiwayatPangkat.Columns(6).Value
    If IsNull(dgRiwayatPangkat.Columns(8).Value) Then txtTandaTanganSK.Text = "" Else txtTandaTanganSK.Text = dgRiwayatPangkat.Columns(8).Value
    If IsNull(dgRiwayatPangkat.Columns(9).Value) Then txtKeterangan.Text = "" Else txtKeterangan.Text = dgRiwayatPangkat.Columns(9).Value
    If IsNull(dgRiwayatPangkat.Columns(7).Value) Then dtpTMT.Value = Null Else dtpTMT.Value = dgRiwayatPangkat.Columns(7).Value
    dgNamaPangkat.Visible = False
    
    strSQL = "SELECT Jenis, Kode, PathFile FROM PathFileUpload where Jenis='Pangkat' and Kode='" & txtKdPangkat.Text & "'"
    Call msubRecFO(rs, strSQL)
    lstFile.clear
    If rs.RecordCount <> 0 Then
        For i = 0 To rs.RecordCount - 1
            'txtpathfile = Replace(rs(2), "/", "\")
            lstFile.AddItem Replace(rs(2), "/", "\")
            rs.MoveNext
        Next
    Else
'        txtpathfile.Text = ""
    End If
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpTMT.SetFocus
End Sub

Private Sub dtpTMT_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtTandaTanganSK.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadRiwayatPangkat
End Sub

Private Sub subLoadRiwayatPangkat()
    On Error GoTo errLoad
    strLSQL = "SELECT IdPegawai, NoUrut, KdPangkat, NamaPangkat, NamaGolongan, NoSK, TglSK, TMT, TandaTanganSk, Keterangan, NamaUser FROM v_RiwayatPangkat WHERE IdPegawai = '" & mstrIdPegawai & "' ORDER BY NoUrut"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgRiwayatPangkat
        Set .DataSource = rs
        .Columns("IdPegawai").Width = 0
        .Columns("NoUrut").Width = 1000
        .Columns("NoUrut").Caption = "No. Urut"
        .Columns("KdPangkat").Width = 1000
        .Columns("KdPangkat").Caption = "Kode"
        .Columns("NamaPangkat").Width = 2000
        .Columns("NamaPangkat").Caption = "Pangkat"
        .Columns("NamaGolongan").Width = 1000
        .Columns("NamaGolongan").Caption = "Gol"
        .Columns("NoSK").Width = 2000
        .Columns("NoSK").Caption = "No. SK"
        .Columns("TglSK").Width = 1500
        .Columns("TglSK").Caption = "Tgl. SK"
        .Columns("TandaTanganSK").Width = 1700
        .Columns("TandaTanganSK").Caption = "TTD SK"
        .Columns("Keterangan").Width = 3700
        .Columns("NamaUser").Caption = "Nama User"
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subClearData()
    On Error Resume Next
    txtNoUrut.Text = ""
    txtKdPangkat.Text = ""
    txtNamaPangkat.Text = ""
    txtNoSK.Text = ""
    dtpTglSK.Value = Format(Now, "dd/mm/yyyy")
    txtTandaTanganSK.Text = ""
    txtKeterangan.Text = ""
    dgNamaPangkat.Visible = False
    dtpTMT.Value = Format(Now, "dd/mm/yyyy")
    txtGol.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call frmRiwayatPegawai.subLoadRiwayatPangkat
    frmRiwayatPegawai.Enabled = True
End Sub

Private Sub Text1_Change()

End Sub



Private Sub lstFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
Dim fso As New FileSystemObject
Dim anjing As Integer
        
    For anjing = 1 To 50
        If fso.FileExists(Data.Files(anjing)) Then lstFile.AddItem Data.Files(anjing)
    Next
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNamaPangkat_Change()
    Call subLoadGridNamaPangkat
End Sub

Private Sub txtNamaPangkat_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgNamaPangkat.Visible = True Then dgNamaPangkat.SetFocus
End Sub

Private Sub txtNamaPangkat_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then If dgNamaPangkat.Visible = True Then dgNamaPangkat.SetFocus Else txtNoSK.SetFocus
End Sub

Private Sub txtNoSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglSK.SetFocus
End Sub

'Private Sub txtpathfile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
'Dim fso As New FileSystemObject
'Dim anjing As Integer
'
'    For anjing = 1 To 10
'        If fso.FileExists(Data.Files(anjing)) Then txtpathfile.Text = txtpathfile.Text & "|" & Data.Files(anjing)
'    Next
'
'End Sub

Private Sub txtTandaTanganSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

