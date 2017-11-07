VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10a.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPenilaianPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Pelayanan Tindakan"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPenilaianPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11775
   Begin MSComctlLib.ListView lvPemeriksa 
      Height          =   1815
      Left            =   4080
      TabIndex        =   23
      Top             =   -840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nama Pemeriksa"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fraPelayanan 
      Caption         =   "Data Pelayanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      TabIndex        =   19
      Top             =   3720
      Visible         =   0   'False
      Width           =   9855
      Begin MSDataGridLib.DataGrid dgPelayanan 
         Height          =   2415
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4260
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
   End
   Begin VB.Frame fraDokter 
      Caption         =   "Data Dokter Pemeriksa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   480
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   8895
      Begin MSDataGridLib.DataGrid dgDokter 
         Height          =   2295
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4048
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
   End
   Begin VB.Frame fradoa 
      Caption         =   "Daftar Layanan Obat && Alkes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   21
      Top             =   5880
      Width           =   9855
      Begin MSFlexGridLib.MSFlexGrid fgDOA 
         Height          =   1335
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   50
         Cols            =   10
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   8577768
         ForeColorFixed  =   -2147483627
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "Daftar Layanan Tindakan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   0
      TabIndex        =   25
      Top             =   3840
      Width           =   9855
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgPelayanan 
         Height          =   1575
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   2778
         _Version        =   393216
         Rows            =   50
         Cols            =   5
         FixedCols       =   0
         BackColorFixed  =   8577768
         BackColorBkg    =   16777215
         FocusRect       =   0
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame fraButton 
      Enabled         =   0   'False
      Height          =   735
      Left            =   0
      TabIndex        =   27
      Top             =   3000
      Width           =   11295
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "&Simpan"
         Height          =   360
         Left            =   8880
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Tutu&p"
         Height          =   360
         Left            =   10080
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraPPelayanan 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   28
      Top             =   2160
      Width           =   11295
      Begin VB.OptionButton optNonPaket 
         Caption         =   " Non Paket"
         Height          =   375
         Left            =   5880
         TabIndex        =   12
         Top             =   550
         Width           =   1215
      End
      Begin VB.OptionButton optPaket 
         Caption         =   " Paket"
         Height          =   375
         Left            =   5880
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtKuantitas 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   5040
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtNamaPelayanan 
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
         Height          =   330
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   4695
      End
      Begin VB.CheckBox chkAPBD 
         Caption         =   "Pos APBD"
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   518
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&Tambah"
         Height          =   360
         Left            =   8760
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "&Hapus"
         Height          =   360
         Left            =   9960
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah"
         Height          =   240
         Left            =   5040
         TabIndex        =   30
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pelayanan"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame fraPDokter 
      Height          =   1095
      Left            =   0
      TabIndex        =   31
      Top             =   1080
      Width           =   11775
      Begin VB.TextBox txtDokter2 
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
         Height          =   330
         Left            =   4840
         TabIndex        =   35
         Top             =   525
         Width           =   2655
      End
      Begin VB.CheckBox chkDelegasi 
         Caption         =   "Di Delegasikan"
         Height          =   255
         Left            =   4800
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Status CITO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9960
         TabIndex        =   32
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optCito 
            Caption         =   "Ya"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optCito 
            Caption         =   "Tidak"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CheckBox chkPerawat 
         Caption         =   "Paramedis"
         Height          =   255
         Left            =   7600
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDokter 
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
         Height          =   330
         Left            =   2160
         TabIndex        =   3
         Top             =   525
         Width           =   2655
      End
      Begin VB.TextBox txtNamaPerawat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7600
         TabIndex        =   5
         Text            =   "txtNamaPerawat"
         Top             =   525
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtpTglPeriksa 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   525
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy HH:mm"
         Format          =   65077251
         UpDown          =   -1  'True
         CurrentDate     =   37823
      End
      Begin VB.CheckBox chkDilayaniDokter 
         Caption         =   "Dokter Pemeriksa "
         Height          =   255
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Periksa"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1365
      End
   End
   Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
      Height          =   1215
      Left            =   5400
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
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
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmPenilaianPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9960
      Picture         =   "frmPenilaianPegawai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPenilaianPegawai.frx":4B79
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmPenilaianPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strFilterPelayanan As String
Dim strCito As String
Dim strKodePelayananRS As String
Dim curBiaya As Currency
Dim curJP As Currency
Dim intJmlPelayanan As Integer
Dim strKdKelas As String
Dim strKelas As String
Dim strKdJenisTarif As String
Dim strJenisTarif As String
Dim intBarang As Integer
Dim intJmlBarang As Integer
Dim intMaxJmlBarang As Integer
Dim strStatusAPBD As String

Dim subKdPemeriksa() As String
Dim subJmlTotal As Integer
Dim curTarifCito As Currency
Dim subcurTarifCito As Currency
Dim subcurTarifBiayaSatuan As Currency
Dim subcurTarifHargaSatuan As Currency
Dim mstrKdDokter2 As String
Dim strPilihGrid As String
Dim i As Integer
Dim j As Integer


Private Function sp_DelegasiBiayaPelayanan(f_NoPendaftaran As String, f_KdRuangan As String, f_KdPelayananRS As String, f_tglPelayanan As Date, f_IdDokterDelegasi As String) As Boolean
On Error GoTo errLoad

    sp_DelegasiBiayaPelayanan = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, f_NoPendaftaran)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, f_KdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, f_KdPelayananRS)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(f_tglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("StatusDelegasi", adChar, adParamInput, 1, "Y")
        .Parameters.Append .CreateParameter("IdDokterDelegasi", adChar, adParamInput, 10, IIf(f_IdDokterDelegasi = "", Null, f_IdDokterDelegasi))
    
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_DelegasiBiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            sp_DelegasiBiayaPelayanan = False
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
        
        End If
    End With
    
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    
Exit Function
errLoad:
    sp_DelegasiBiayaPelayanan = False
    Call msubPesanError("sp_DelegasiBiayaPelayanan")
End Function

Private Sub chkAPBD_Click()
    If chkAPBD.Value = 1 Then
        strStatusAPBD = "01"
    Else
        strStatusAPBD = "02"
    End If
End Sub

Private Sub chkAPBD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPelayanan.SetFocus
End Sub

Private Sub chkDelegasi_Click()
If chkDelegasi.Value = vbChecked Then
If MsgBox("Akan Didelegasikan Ke Dokter Atau Paramedis ?? " & vbCrLf & "Pilih YES Untuk DOKTER atau Pilih NO Untuk PARAMEDIS ", vbYesNo, "Validasi") = vbYes Then
    chkPerawat.Value = vbUnchecked
    chkPerawat.Enabled = False
    txtDokter2.Enabled = True
    lvPemeriksa.Enabled = False
Else
    chkPerawat.Value = vbChecked
    chkPerawat.Enabled = True
    txtDokter2.Enabled = False
    lvPemeriksa.Enabled = True
End If
Else
    chkPerawat.Value = vbChecked
    chkPerawat.Enabled = True
    txtDokter2.Enabled = False
    lvPemeriksa.Enabled = True
End If
End Sub

Private Sub chkDelegasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkPerawat.SetFocus
End Sub

Private Sub chkDilayaniDokter_Click()
On Error GoTo errLoad
    
    If chkDilayaniDokter.Value = 0 Then
        txtDokter.Enabled = False
        txtDokter.Text = ""
        
        If fraDokter.Visible = True Then fraDokter.Visible = False
    Else
        lvPemeriksa.Visible = False
        
        txtDokter.Enabled = True
        strSQL = "SELECT dbo.RegistrasiRI.IdDokter, dbo.DataPegawai.NamaLengkap " & _
            " FROM dbo.RegistrasiRI INNER JOIN dbo.DataPegawai ON dbo.RegistrasiRI.IdDokter = dbo.DataPegawai.IdPegawai " & _
            " WHERE (dbo.RegistrasiRI.NoPendaftaran = '" & mstrNoPen & "')"
        Call msubRecFO(rs, strSQL)
        
        If Not rs.EOF Then
            txtDokter.Text = rs(1).Value
            mstrKdDokter = rs(0).Value
            intJmlDokter = rs.RecordCount
            fraDokter.Visible = False
        End If
    End If
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub chkDilayaniDokter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkDilayaniDokter.Value = 0 Then
            chkPerawat.SetFocus
        Else
            txtDokter.SetFocus
        End If
    End If
End Sub

Private Sub chkPerawat_Click()
    If chkPerawat.Value = vbChecked Then
        strSQL = "SELECT IdPegawai FROM V_DaftarPemeriksaPasien WHERE (IdPegawai = '" & strIDPegawaiAktif & "')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = False Then
            txtNamaPerawat.Text = strNmPegawai
            If lvPemeriksa.ListItems.Count > 0 Then
                lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif).Checked = True
                Call lvPemeriksa_ItemCheck(lvPemeriksa.ListItems.Item("key" & strIDPegawaiAktif))
            End If
        Else
            txtNamaPerawat.Text = ""
        End If
    Else
        txtNamaPerawat.Text = ""
    End If
    lvPemeriksa.Visible = False
End Sub

Private Sub chkPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkPerawat.Value = vbChecked Then
            txtNamaPerawat.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
End Sub

Private Sub cmdBatal_Click()
If txtNamaPelayanan.Text = "" Then Unload Me: Exit Sub
    If cmdSimpan.Enabled = True Then
        If MsgBox("Simpan data tindakan pasien?", vbQuestion + vbYesNo, "Konfirmasi") = vbYes Then
            Call cmdSimpan_Click
            Exit Sub
        End If
    End If
    Unload Me
    frmTransaksiPasien.Enabled = True
End Sub

Private Sub cmdHapus_Click()
Dim h As Integer
Dim j As Integer
    With fgPelayanan
        If .Row = .Rows Then Exit Sub
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        h = 1
        Do While h <= fgDOA.Rows - 2
            If fgDOA.TextMatrix(h, 9) = .TextMatrix(.Row, 0) Then
'-----------------------**--Yang ditambah--**-----------------------
                For j = 1 To intMaxJmlBarang
                    If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then
                        If fgDOA.TextMatrix(h, 5) = "S" Then
                            typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + (fgDOA.TextMatrix(h, 3) * typBarang(j).intJmlTerkecil)
                        Else
                            typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + fgDOA.TextMatrix(h, 3)
                        End If
                    End If
                Next j
'-----------------------**--Yang ditambah--**-----------------------
'                fgDOA.RemoveItem h
                Call msubRemoveItem(fgDOA, h)
                h = 0
            End If
            h = h + 1
        Loop
'-----------------------**--Yang ditambah--**-----------------------
        For j = 1 To intMaxJmlBarang
            For h = 1 To fgDOA.Rows - 1
                If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then Exit For
                If h = fgDOA.Rows - 1 Then
                    intMaxJmlBarang = intMaxJmlBarang - 1
                    If intMaxJmlBarang < 0 Then intMaxJmlBarang = 0
                End If
            Next h
        Next j
'-----------------------**--Yang ditambah--**-----------------------
'        .RemoveItem .Row
        Call msubRemoveItem(fgPelayanan, .Row)
    End With
End Sub

'Store procedure untuk mengisi registrasi pasien
Private Sub sp_UbahDokter(ByVal adoCommand As ADODB.Command)
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, mstrKdDokter)
        .Parameters.Append .CreateParameter("TglMasuk", adDate, adParamInput, , Format(dtpTglPeriksa, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 10, mstrKdRuangan)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Update_DokterPemeriksaRISewaKamar"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam penyimpanan Dokter Pemeriksa pasien", vbCritical, "Validasi"
        Else
'            MsgBox "Penyimpanan Dokter Pemeriksa pasien sukses", vbInformation, "Informasi"
            
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdSimpan_Click()
On Error GoTo errLoad
Dim tempKdRuanganAsal As String
    
    If funcCekValidasi = False Then Exit Sub
    Call subEnableButtonReg(False)
    For i = 1 To fgPelayanan.Rows - 2
        'simpan biaya pelayanan
        If sp_BiayaPelayanan(dbcmd, fgPelayanan.TextMatrix(i, 0), CCur(fgPelayanan.TextMatrix(i, 3)), fgPelayanan.TextMatrix(i, 2), fgPelayanan.TextMatrix(i, 9), fgPelayanan.TextMatrix(i, 6), fgPelayanan.TextMatrix(i, 7), CCur(fgPelayanan.TextMatrix(i, 8))) = False Then Exit Sub
        
        Set rs = Nothing
        If txtDokter.Text = "" Then GoTo skipp_
        strSQL = "select KdJnsPelayanan from ListPelayananRS where KdPelayananRS= '" & fgPelayanan.TextMatrix(i, 0) & "'"
        Call msubRecFO(rs, strSQL)
        If rs.Fields("KdJnsPelayanan") = "303" Or rs.Fields("KdJnsPelayanan") = "305" Then
            Call sp_UbahDokter(dbcmd)
        End If
        
'        'ambil kdruangan asal
'        Call msubRecFO(rs, "SELECT dbo.FB_TakeRuanganAsal('" & mstrNoPen & "', '" & mstrKdRuangan & "', null,'" & Format(fgPelayanan.TextMatrix(i, 9), "yyyy/MM/dd HH:mm:ss") & "','TM') as KdRuanganAsal")
'        If rs.EOF = False Then tempKdRuanganAsal = rs(0) Else tempKdRuanganAsal = ""
'        'simpan temp harga komponen
'        If functAdd_TempHargaKomponen(mstrNoPen, mstrKdRuangan, fgPelayanan.TextMatrix(i, 9), fgPelayanan.TextMatrix(i, 0), strKdKelas, strKdJenisTarif, CCur(fgPelayanan.TextMatrix(i, 8)), fgPelayanan.TextMatrix(i, 2), fgPelayanan.TextMatrix(i, 7), strIDPegawaiAktif, tempKdRuanganAsal) = False Then Exit Sub

        'simpan delegasi biaya pelayanan jika status 'Y'
skipp_:
          If chkDelegasi.Value = vbChecked Then If sp_DelegasiBiayaPelayanan(mstrNoPen, mstrKdRuangan, fgPelayanan.TextMatrix(i, 0), fgPelayanan.TextMatrix(i, 9), fgPelayanan.TextMatrix(i, 10)) = False Then Exit Sub
        
    Next i
    
    If chkPerawat.Value = Checked Then
        For i = 1 To fgPerawatPerPelayanan.Rows - 1
            With fgPerawatPerPelayanan
                If sp_PetugasPemeriksaBP(.TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), .TextMatrix(i, 5)) = False Then Exit Sub
            End With
        Next i
    End If
    
Dim adoCommand As New ADODB.Command
    If fgDOA.Rows = 2 Then GoTo stepNonPaketSemua
    For i = 1 To fgDOA.Rows - 2
        With adoCommand
            .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
            .Parameters.Append .CreateParameter("KdBarang", adVarChar, adParamInput, 9, fgDOA.TextMatrix(i, 0))
            .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 2, fgDOA.TextMatrix(i, 2))
            .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
            .Parameters.Append .CreateParameter("Satuan", adChar, adParamInput, 1, fgDOA.TextMatrix(i, 5))
            .Parameters.Append .CreateParameter("JmlBrg", adInteger, adParamInput, , fgDOA.TextMatrix(i, 3))
            .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
            .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
            .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
            .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(fgDOA.TextMatrix(i, 4)))
            .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(fgDOA.TextMatrix(i, 7), "yyyy/MM/dd HH:mm:ss"))
            .Parameters.Append .CreateParameter("NoLabRad", adChar, adParamInput, 10, Null)
            .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, fgDOA.TextMatrix(i, 6))
            .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
            .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)
            
            .ActiveConnection = dbConn
            .CommandText = "dbo.Add_PemakaianObatAlkes"
            .CommandType = adCmdStoredProc
            .Execute
            
            If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                MsgBox "Ada Kesalahan dalam Penyimpanan Paket Pelayanan Pasien", vbCritical, "Validasi"
'-----------------------**--Yang ditambah--**-----------------------
                Call deleteADOCommandParameters(adoCommand)
                Set adoCommand = Nothing
                GoTo stepErrorPaket
'-----------------------**--Yang ditambah--**-----------------------
            
            End If
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
        End With
    Next i
    
    Call Add_HistoryLoginActivity("Add_BiayaPelayanan+Update_DokterPemeriksaRISewaKamar+Add_DelegasiBiayaPelayanan+Add_PetugasPemeriksaBP+Add_PemakaianObatAlkes")
stepNonPaketSemua:
'    MsgBox "Pemasukan Biaya Pelayanan Pasien Sukses", vbInformation, "Informasi"
'-----------------------**--Yang ditambah--**-----------------------
stepErrorPaket:
'-----------------------**--Yang ditambah--**-----------------------
    frmTransaksiPasien.subLoadPelayananDidapat
    frmTransaksiPasien.subPemakaianObatAlkes
    frmTransaksiPasien.subLoadRiwayatPemeriksaan False
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

'Store procedure untuk menghapus biaya pelayanan pasien yang gagal disimpan
Private Sub sp_DelBiayaPelayananCek(varKdPelayananRS As String, varTglPelayanan As Date)
Dim adoCek As ADODB.Command
    Set adoCek = New ADODB.Command
    With adoCek
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, varKdPelayananRS)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(varTglPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Delete_BiayaPelayananNew"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penghapusan Biaya Pelayanan Pasien", vbCritical, "Validasi"
        Else
'            MsgBox "Pemasukan Biaya Pelayanan Pasien sukses", vbExclamation, "Validasi"
            Call Add_HistoryLoginActivity("Delete_BiayaPelayananNew")
        End If
        Call deleteADOCommandParameters(adoCek)
        Set adoCek = Nothing
    End With
    Exit Sub
End Sub

Private Sub cmdTambah_Click()
Dim i As Integer
Dim j As Integer
Dim h As Integer
Dim adocmd As New ADODB.Command
On Error Resume Next
    If chkDilayaniDokter.Value = vbChecked Then
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter Pemeriksa Pasien", vbCritical, "Validasi"
            txtDokter.SetFocus
            Exit Sub
        End If
    End If
    If chkDelegasi.Value = vbChecked And chkPerawat.Value = vbUnchecked Then
        If txtDokter2.Text = "" Then
            MsgBox "Pilih dulu Dokter yg didelegasikan!!", vbCritical, "Validasi"
            txtDokter2.SetFocus
            Exit Sub
        End If
    End If
    If chkPerawat.Value = vbChecked And subJmlTotal = 0 Then
        MsgBox "Nama perawat kosong", vbCritical, "Validasi"
        lvPemeriksa.Visible = True
        txtNamaPerawat.SetFocus
        Exit Sub
    End If
    
    If strKodePelayananRS = "" Then Exit Sub
    If optNonPaket.Value = True Then GoTo stepNonPaket
Dim dTglPlyn As Date
    dTglPlyn = Now
    strSQL = "Select * FROM V_PaketPelayananObatAlkes WHERE KdPelayananRS='" & strKodePelayananRS & "' AND KdKelompokPasien = '" & mstrKdJenisPasien & "' AND IdPenjamin = '" & mstrKdPenjaminPasien & "'"
    Call msubRecFO(rs, strSQL)
'-----------------------**--Yang ditambah--**-----------------------
    For i = 1 To rs.RecordCount
        'cek data barang & asal barang di grid paket obat alkes
        For j = 1 To fgDOA.Rows - 1
            'barang dengan asal barang tersebut sudah ada di grid obat alkes
            If fgDOA.TextMatrix(j, 0) = rs("KdBarang").Value And fgDOA.TextMatrix(j, 2) = rs("KdAsal").Value Then
                For h = 1 To intMaxJmlBarang
                    If typBarang(h).strkdbarang = rs("KdBarang").Value And typBarang(h).strkdasal = rs("KdAsal").Value Then
                        intJmlBarang = h
                        GoTo stepCekStokBarang
                    End If
                Next h
            End If
            'sampai data terakhir data barang tidak ada di grid obat alkes
            If j = fgDOA.Rows - 1 Then
                'tambahkan data total barang yang terpakai
                intMaxJmlBarang = intMaxJmlBarang + 1
                intJmlBarang = intMaxJmlBarang
ReDim Preserve typBarang(intMaxJmlBarang)
                strSQL = "SELECT JmlTerkecil,JmlTotalBarangTemp,NamaBarang FROM " _
                    & "V_StokBarangTempRuangan WHERE KdBarang='" _
                    & rs("KdBarang").Value & "' AND KdAsal='" _
                    & rs("KdAsal").Value & "' AND KdRuangan='" _
                    & mstrKdRuangan & "'"
                Set rsb = Nothing
                rsb.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
                typBarang(intJmlBarang).strkdbarang = rs("KdBarang").Value
                typBarang(intJmlBarang).strNamaBarang = rsb("NamaBarang").Value
                typBarang(intJmlBarang).strkdasal = rs("KdAsal").Value
                typBarang(intJmlBarang).intJmlTerkecil = rsb("JmlTerkecil").Value
                typBarang(intJmlBarang).intJmlTempTotal = rsb("JmlTotalBarangTemp").Value
            End If
        Next j
stepCekStokBarang:
        If funcCekStokBarang(intJmlBarang, rs("SatuanJml"), (CInt(txtKuantitas) * rs("JmlBarang").Value)) = False Then
            'hapus grid obat alkes dengan kode pelayanan tersebut
            h = 1
            Do While h <= fgDOA.Rows - 2
                If fgDOA.TextMatrix(h, 9) = strKodePelayananRS Then
                    For j = 1 To intMaxJmlBarang
                        If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then
                            If fgDOA.TextMatrix(h, 5) = "S" Then
                                typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + (fgDOA.TextMatrix(h, 3) * typBarang(j).intJmlTerkecil)
                            Else
                                typBarang(j).intJmlTempTotal = typBarang(j).intJmlTempTotal + fgDOA.TextMatrix(h, 3)
                            End If
                        End If
                    Next j
                    fgDOA.RemoveItem h
                    h = 0
                End If
                h = h + 1
            Loop
            h = 1
            For j = 1 To intMaxJmlBarang
                For h = 1 To fgDOA.Rows - 1
                    If typBarang(j).strkdbarang = fgDOA.TextMatrix(h, 0) Then Exit For
                    If h = fgDOA.Rows - 1 Then
                        intMaxJmlBarang = intMaxJmlBarang - 1
                        If intMaxJmlBarang < 0 Then intMaxJmlBarang = 0
                    End If
                Next h
            Next j
            Exit Sub
        End If
'-----------------------**--Yang ditambah--**-----------------------
        With fgDOA
            mintRowNow = .Rows - 1
            .TextMatrix(mintRowNow, 0) = rs("KdBarang").Value
            .TextMatrix(mintRowNow, 1) = rs("NamaBarang").Value
            .TextMatrix(mintRowNow, 2) = rs("KdAsal").Value
            .TextMatrix(mintRowNow, 3) = CInt(txtKuantitas) * rs("JmlBarang").Value
            
'            subcurTarifHargaSatuan = sp_Take_TarifOA(rs("KdAsal").Value, rs("Harga").Value)
'            .TextMatrix(mintRowNow, 4) = subcurTarifHargaSatuan 'rs("Harga").Value
            
            strSQL = "SELECT dbo.FB_TakeTarifOA('" & mstrKdJenisPasien & "','" & mstrKdPenjaminPasien & "','" & rs("KdAsal").Value & "', " & CCur(rs("HargaBarang").Value) & ")  as HargaSatuan"
            Call msubRecFO(dbRst, strSQL)
            If dbRst.EOF = True Then subcurTarifHargaSatuan = 0 Else subcurTarifHargaSatuan = dbRst(0).Value
            .TextMatrix(mintRowNow, 4) = subcurTarifHargaSatuan
            
            .TextMatrix(mintRowNow, 5) = rs("SatuanJml").Value
            If chkDilayaniDokter.Value = 1 Then
                .TextMatrix(mintRowNow, 6) = mstrKdDokter
            Else
                .TextMatrix(mintRowNow, 6) = UserID
            End If
            .TextMatrix(mintRowNow, 7) = Format(dTglPlyn, "dd/mm/yyyy HH:mm:ss")
            .TextMatrix(mintRowNow, 8) = rs("NamaAsal").Value
            .TextMatrix(mintRowNow, 9) = strKodePelayananRS
            .Rows = .Rows + 1
            .SetFocus
        End With
        rs.MoveNext
    Next i
stepNonPaket:
    With fgPelayanan
        For i = 1 To .Rows - 1
            If (.TextMatrix(i, 0) = strKodePelayananRS) And _
               (.TextMatrix(i, 9) = dtpTglPeriksa.Value) Then txtNamaPelayanan.SetFocus: txtNamaPelayanan.SelStart = 0: txtNamaPelayanan.SelLength = Len(txtNamaPelayanan.Text): Exit Sub
        Next i
        intRowNow = .Rows - 1
        .TextMatrix(intRowNow, 0) = strKodePelayananRS
        .TextMatrix(intRowNow, 1) = txtNamaPelayanan.Text
        .TextMatrix(intRowNow, 2) = CInt(txtKuantitas.Text)
        
        subcurTarifCito = sp_Take_TarifBPT
        .TextMatrix(intRowNow, 3) = IIf(subcurTarifBiayaSatuan = 0, 0, Format(subcurTarifBiayaSatuan, "#,###")) 'curBiaya
        .TextMatrix(intRowNow, 4) = IIf(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text) = 0, 0, Format(funcRoundUp(CStr(subcurTarifBiayaSatuan + subcurTarifCito)) * CInt(txtKuantitas.Text), "#,###"))
        .TextMatrix(intRowNow, 8) = subcurTarifCito
        
        .TextMatrix(intRowNow, 5) = mdTglBerlaku
        If chkDilayaniDokter.Value = 1 Then
            .TextMatrix(intRowNow, 6) = mstrKdDokter
        Else
            .TextMatrix(intRowNow, 6) = UserID
        End If
        .TextMatrix(intRowNow, 7) = strCito
        .TextMatrix(intRowNow, 9) = dtpTglPeriksa.Value
       If chkDelegasi.Value = vbChecked And txtDokter2.Text = "" Then
            .TextMatrix(intRowNow, 10) = ""
        ElseIf chkDelegasi.Value = vbChecked And txtDokter2.Text <> "" Then
            .TextMatrix(intRowNow, 10) = mstrKdDokter2
        End If
        
        .Rows = .Rows + 1
        .SetFocus
    End With
    
    If chkPerawat.Value = vbChecked Then Call subLoadPelayananPerPerawat
    txtNamaPelayanan.Text = ""
    txtKuantitas.Text = 1
    fraPelayanan.Visible = False
    chkPerawat.SetFocus
End Sub

Private Sub subLoadPelayananPerPerawat()
    With fgPerawatPerPelayanan
        For i = 1 To subJmlTotal
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = mstrNoPen
            .TextMatrix(.Rows - 1, 1) = mstrKdRuangan
            .TextMatrix(.Rows - 1, 2) = dtpTglPeriksa.Value
            .TextMatrix(.Rows - 1, 3) = strKodePelayananRS
            .TextMatrix(.Rows - 1, 4) = Mid(subKdPemeriksa(i), 4, Len(subKdPemeriksa(i)) - 3)
            .TextMatrix(.Rows - 1, 5) = strIDPegawaiAktif
        Next
    End With

    subJmlTotal = 0
    txtNamaPerawat.BackColor = &HFFFFFF
    ReDim Preserve subKdPemeriksa(subJmlTotal)
   ' chkPerawat.Value = vbUnchecked
End Sub

Private Sub dgDokter_DblClick()
    Call dgDokter_KeyPress(13)
End Sub

Private Sub dgDokter_KeyPress(KeyAscii As Integer)
On Error Resume Next
If strPilihGrid = "Dokter" Then
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokter.Text = dgDokter.Columns(1).Value
        mstrKdDokter = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        chkDilayaniDokter.Value = 1
        fraDokter.Visible = False
        chkPerawat.SetFocus
    End If
ElseIf strPilihGrid = "Dokter2" Then
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        txtDokter2.Text = dgDokter.Columns(1).Value
        mstrKdDokter2 = dgDokter.Columns(0).Value
        If mstrKdDokter = "" Then
            MsgBox "Pilih dulu Dokter yang akan menangani Pasien", vbCritical, "Validasi"
            txtDokter2.Text = ""
            dgDokter.SetFocus
            Exit Sub
        End If
        fraDokter.Visible = False
        chkPerawat.SetFocus
    End If
If KeyAscii = 27 Then
    fraDokter.Visible = False
End If
End If
End Sub

Private Sub dgPelayanan_DblClick()
    Call dgPelayanan_KeyPress(13)
End Sub

Private Sub dgPelayanan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
Dim strkd As String
        strkd = dgPelayanan.Columns(5).Value
        curBiaya = dgPelayanan.Columns(4).Value
'        curJP = dgPelayanan.Columns(6).Value
        txtNamaPelayanan.Text = dgPelayanan.Columns(1).Value
        strKodePelayananRS = strkd
        optNonPaket.Value = True
        If strKodePelayananRS = "" Then
            MsgBox "Pilih dulu tindakan pelayanan Pasien", vbCritical, "Validasi"
            txtNamaPelayanan.Text = ""
            dgPelayanan.SetFocus
            Exit Sub
        End If
        fraPelayanan.Visible = False
        txtKuantitas.SetFocus
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
End Sub

Private Sub dtpTglPeriksa_Change()
    dtpTglPeriksa.MaxDate = Now
    If dtpTglPeriksa.Value < mdTglMasuk Then dtpTglPeriksa.Value = mdTglMasuk
End Sub

Private Sub dtpTglPeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtDokter.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    frmTransaksiPasien.Enabled = False
'    strSQL = "SELECT KelasPelayanan.KdKelas,KelasPelayanan.DeskKelas " _
        & "FROM RegistrasiRI INNER JOIN KelasPelayanan ON " _
        & "RegistrasiRI.KdKelas = KelasPelayanan.KdKelas " _
        & "WHERE RegistrasiRI.NoPendaftaran='" & mstrNoPen & "'"
'    Set rs = Nothing
'    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockOptimistic
'    strKdKelas = rs.Fields(0).Value
'    strKelas = rs.Fields(1).Value

    strKdKelas = mstrKdKelas
    Set rs = Nothing
    strSQL = "SELECT KdJenisTarif,JenisTarif " _
        & "FROM v_JenisTarifPasien " _
        & "WHERE NoPendaftaran='" & mstrNoPen & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockOptimistic
    strKdJenisTarif = rs.Fields(0).Value
    strJenisTarif = rs.Fields(1).Value
    Set rs = Nothing
    Call subSetGidPelayanan
    dtpTglPeriksa.Value = Now
    strCito = "0"
    strStatusAPBD = "01"
    optNonPaket.Value = True
    Call subSetGridObatAlkes
    
    intBarang = 0
    intJmlBarang = 0
    intMaxJmlBarang = 0
    ReDim typBarang(0)
    
    subJmlTotal = 0
    Call subSetGridPerawatPerPelayanan
    Call subLoadListPemeriksa
    chkPerawat.Value = vbChecked
    lvPemeriksa.Visible = False
    
Exit Sub
errLoad:
    Call msubPesanError
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmTransaksiPasien.Enabled = True
    frmTransaksiPasien.fraDokterP.Visible = False
End Sub

Private Sub lvPemeriksa_DblClick()
    Call lvPemeriksa_KeyPress(13)
End Sub

Private Sub lvPemeriksa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim blnSelected As Boolean
    If Item.Checked = True Then
        subJmlTotal = subJmlTotal + 1
        ReDim Preserve subKdPemeriksa(subJmlTotal)
        subKdPemeriksa(subJmlTotal) = Item.key
    Else
        blnSelected = False
        For i = 1 To subJmlTotal
            If subKdPemeriksa(i) = Item.key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotal Then
                    subKdPemeriksa(i) = ""
                Else
                    subKdPemeriksa(i) = subKdPemeriksa(i + 1)
                End If
            End If
        Next i
        subJmlTotal = subJmlTotal - 1
    End If
    
    If subJmlTotal = 0 Then
        txtNamaPerawat.BackColor = &HFFFFFF
        chkPerawat.Caption = "Paramedis"
    Else
        txtNamaPerawat.BackColor = &HC0FFFF
        chkPerawat.Caption = "Paramedis (" & subJmlTotal & " org)"
    End If
End Sub

Private Sub lvPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvPemeriksa.Visible = False: txtNamaPerawat.SetFocus
End Sub

Private Sub optCito_Click(Index As Integer)
    If Index = 0 Then
        strCito = "1"
    Else
        strCito = "0"
    End If
'    Call txtNamaPelayanan_Change
End Sub

Private Sub optCito_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkAPBD.Enabled = True Then
            chkAPBD.SetFocus
        Else
            txtNamaPelayanan.SetFocus
        End If
    End If
End Sub

Private Sub optNonPaket_Click()
'    fradoa.Visible = False
    fraButton.Enabled = True
End Sub

Private Sub optNonPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub optPaket_Click()
    strSQL = "SELECT * FROM PaketLayanan WHERE KdPelayananRS='" & strKodePelayananRS _
        & "' AND KdRuangan='" & mstrKdRuangan & "'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        MsgBox "Tidak ada paket untuk pelayanan yang dipilih", vbCritical, "Validasi"
'        fradoa.Visible = False
        optNonPaket.SetFocus
    Else
'        fradoa.Visible = True
    End If
    fraButton.Enabled = True
End Sub

Private Sub optPaket_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdTambah.SetFocus
End Sub

Private Sub txtDokter_Change()
    strPilihGrid = "Dokter"
    mstrFilterDokter = "WHERE NamaDokter like '%" & txtDokter.Text & "%'"
    mstrKdDokter = ""
    fraDokter.Visible = True
    Call subLoadDokter
End Sub

Private Sub txtDokter_KeyPress(KeyAscii As Integer)
'On Error GoTo hell
    If KeyAscii = 13 Then
        If mintJmlDokter = 0 Then Exit Sub
        If fraDokter.Visible = True Then
            dgDokter.SetFocus
        Else
            chkDelegasi.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraDokter.Visible = False
    End If
'hell:
End Sub

Private Sub txtDokter2_Change()
    strPilihGrid = "Dokter2"
    fraDokter.Visible = True
    Call subLoadDokter2
End Sub

Private Sub txtDokter2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then If fraDokter.Visible = True Then dgDokter.SetFocus
End Sub

Private Sub txtKuantitas_Change()
    If txtKuantitas.Text = "" Or txtKuantitas.Text = 0 Then txtKuantitas.Text = 1
End Sub

Private Sub txtKuantitas_GotFocus()
    txtKuantitas.SelStart = 0
    txtKuantitas.SelLength = Len(txtKuantitas.Text)
End Sub

Private Sub txtKuantitas_KeyPress(KeyAscii As Integer)
    Call SetKeyPressToNumber(KeyAscii)
    If KeyAscii = 13 Then optNonPaket.SetFocus
End Sub

Private Sub txtKuantitas_LostFocus()
    If txtKuantitas.Text = "" Then txtKuantitas.Text = 1: Exit Sub
    If txtKuantitas.Text = 0 Then txtKuantitas.Text = 1
End Sub

Private Sub txtNamaPelayanan_Change()
    strFilterPelayanan = "WHERE [Nama Pelayanan] like '%" & txtNamaPelayanan.Text _
        & "%' AND KdKelas='" & strKdKelas & "' AND KdJenisTarif='" & strKdJenisTarif _
        & "' AND KdRuangan='" & mstrKdRuangan & "'"
    strKodePelayananRS = ""
    fraPelayanan.Visible = True
    Call subLoadPelayanan
End Sub

Private Sub txtNamaPelayanan_GotFocus()
    lvPemeriksa.Visible = False
End Sub

Private Sub txtNamaPelayanan_KeyPress(KeyAscii As Integer)
On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If intJmlPelayanan = 0 Then Exit Sub
        If fraPelayanan.Visible = True Then
            dgPelayanan.SetFocus
        Else
            txtKuantitas.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        fraPelayanan.Visible = False
    End If
hell:
End Sub

'untuk meload data dokter di grid
Private Sub subLoadDokter()
'    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter " & mstrFilterDokter
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 0
    fraDokter.Top = 1920
End Sub

'untuk meload data dokter delegasi di grid
Private Sub subLoadDokter2()
'    On Error Resume Next
    strSQL = "SELECT KodeDokter AS [Kode Dokter],NamaDokter AS [Nama Dokter],JK,Jabatan FROM V_DaftarDokter WHERE NamaDokter like '%" & txtDokter2.Text & "%'"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    mintJmlDokter = rs.RecordCount
    With dgDokter
        Set .DataSource = rs
        .Columns(0).Width = 1200
        .Columns(1).Width = 3000
        .Columns(2).Width = 400
        .Columns(3).Width = 3000
    End With
    fraDokter.Left = 4000
    fraDokter.Top = 1920
End Sub

'untuk meload data pelayanan di grid
Private Sub subLoadPelayanan()
    On Error Resume Next
    strSQL = "SELECT [Jenis Pelayanan],[Nama Pelayanan],Kelas,JenisTarif,Tarif,KdPelayananRS FROM V_TarifPelayananTindakan " & strFilterPelayanan
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intJmlPelayanan = rs.RecordCount
    With dgPelayanan
        Set .DataSource = rs
        .Columns(0).Width = 2100
        .Columns(1).Width = 3900
        .Columns(2).Width = 1000
        .Columns(3).Width = 1100
        .Columns(4).Width = 900
        .Columns(4).Alignment = dbgRight
        .Columns(5).Width = 0
    End With
    fraPelayanan.Left = 0
    fraPelayanan.Top = 3240
End Sub

'Store procedure untuk mengisi biaya pelayanan pasien
'-----------------------**--Yang ditambah--**-----------------------
Private Function sp_BiayaPelayanan(ByVal adoCommand As ADODB.Command, strKdPelayananRS As String, curTarif As Currency, intJmlPel As Integer, dtTanggalPelayanan As Date, strkodedokter As String, strStatusCITO As String, f_TarifCito As Currency) As Boolean
On Error GoTo errLoad
    sp_BiayaPelayanan = True
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdSubInstalasi", adChar, adParamInput, 3, mstrKdSubInstalasi)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKdPelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, strKdKelas)
        .Parameters.Append .CreateParameter("StatusCITO", adChar, adParamInput, 1, strStatusCITO)
        .Parameters.Append .CreateParameter("Tarif", adInteger, adParamInput, , curTarif)
        .Parameters.Append .CreateParameter("JmlPelayanan", adInteger, adParamInput, , intJmlPel)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        
        Call msubRecFO(rs, "SELECT KdPelayananRS FROM dbo.PelayananRuangan WHERE (Status IN ('CU', 'MA', 'RG')) AND (KdPelayananRS = '" & strKdPelayananRS & "')")
        If rs.EOF = False Then
            Call msubRecFO(rs, "SELECT NoPakai FROM dbo.V_DaftarPasienRIAktif WHERE (NoPendaftaran = '" & mstrNoPen & "')")
            If rs.EOF = False Then
                .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, rs(0))
            Else
                .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
            End If
        Else
            .Parameters.Append .CreateParameter("NoLab_Rad", adChar, adParamInput, 10, Null)
        End If
        
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strkodedokter)
        .Parameters.Append .CreateParameter("StatusAPBD", adChar, adParamInput, 2, strStatusAPBD)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adInteger, adParamInput, , f_TarifCito)
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, Null)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_BiayaPelayanan"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada Kesalahan dalam Penyimpanan Biaya Pelayanan Pasien", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            sp_BiayaPelayanan = False
            GoTo errLoad
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With

Exit Function
errLoad:
        'cek Hasil Entry
        strSQLCari = "select  a.NoPendaftaran, b.kdRuangan, b.TglPelayanan, c.kdPelayananRS, c.kdKomponen from BiayaPelayanan a inner join DetailBiayaPelayanan b on a.NoPendaftaran = b.NoPendaftaran and a.kdPelayananRS = b.kdPelayananRS and a.tglPelayanan = b.tglPelayanan and a.kdRuangan = b.kdRuangan inner join TempHargaKomponen c on a.NoPendaftaran = c.NoPendaftaran and a.kdPelayananRS = c.kdPelayananRS and a.tglPelayanan = c.tglPelayanan and a.kdRuangan = c.kdRuangan " & _
                    " where a.NoPendaftaran = '" & mstrNoPen & "' and b.kdRuangan = '" & mstrKdRuangan & "' and b.kdpelayananrs = '" & strKdPelayananRS & "' " & _
                    " and c.TglPelayanan = '" & Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss") & "' and c.kdKomponen <> '12'"
        Call msubRecFO(rsCari, strSQLCari)
        If rsCari.RecordCount = 0 Then
            sp_DelBiayaPelayananCek strKdPelayananRS, Format(dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss")
            MsgBox "Penyimpanan gagal." & vbCrLf & "Segera hubungi Support System", vbOKOnly + vbCritical, "ADMINISTRATOR"
        End If
        'END cek Hasil Entry
    sp_BiayaPelayanan = False
    Call msubPesanError
End Function

'simpan data perawat
Private Function sp_PetugasPemeriksaBP(F_dtTanggalPelayanan As Date, F_strKodePelayanan As String, F_StrIdPerawat As String, F_IdUser As String) As Boolean
On Error GoTo errLoad

    sp_PetugasPemeriksaBP = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdRuangan", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("TglPelayanan", adDate, adParamInput, , Format(F_dtTanggalPelayanan, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, F_strKodePelayanan)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, F_StrIdPerawat)  'kode perawat
        .Parameters.Append .CreateParameter("IdUser", adChar, adParamInput, 10, F_IdUser)
        
        .ActiveConnection = dbConn
        .CommandText = "dbo.Add_PetugasPemeriksaBP"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data petugas pemeriksa BP", vbExclamation, "Validasi"
            sp_PetugasPemeriksaBP = False
        
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

Exit Function
errLoad:
    sp_PetugasPemeriksaBP = False
    Call msubPesanError
End Function

'untuk set grid pelayanan
Private Sub subSetGidPelayanan()
    With fgPelayanan
        .clear
        .Rows = 2
        .Cols = 11
        .TextMatrix(0, 0) = "Kode Pelayanan"
        .TextMatrix(0, 1) = "Nama Pelayanan"
        .TextMatrix(0, 2) = "Jumlah"
        .TextMatrix(0, 3) = "Biaya Satuan"
        .TextMatrix(0, 4) = "Biaya Total"
        .TextMatrix(0, 5) = "Tgl Berlaku"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "Status CITO"
        .TextMatrix(0, 8) = "Biaya CITO"
        .TextMatrix(0, 9) = "Tgl Pelayanan"
        .TextMatrix(0, 10) = "DokterDelegasi"
        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 700
        .ColWidth(3) = 1200
        .ColWidth(4) = 1400
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1200
        .ColWidth(9) = 0
        .ColWidth(10) = 0
    End With
End Sub

'untuk set grid obat alkes
Private Sub subSetGridObatAlkes()
    With fgDOA
        .clear
        .Rows = 2
        .Cols = 10
        .TextMatrix(0, 0) = "Kode Barang"
        .TextMatrix(0, 1) = "Nama Barang"
        .TextMatrix(0, 2) = "Kode Asal"
        .TextMatrix(0, 3) = "Jumlah"
        .TextMatrix(0, 4) = "Harga Satuan"
        .TextMatrix(0, 5) = "Satuan"
        .TextMatrix(0, 6) = "Kode Dokter"
        .TextMatrix(0, 7) = "tgl Pelayanan"
        .TextMatrix(0, 8) = "Asal Barang"
        .TextMatrix(0, 9) = "KdPelayananRS"
        .ColWidth(0) = 0
        .ColWidth(1) = 4500
        .ColWidth(2) = 0
        .ColWidth(3) = 700
        .ColWidth(4) = 1200
        .ColWidth(5) = 900
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ColWidth(8) = 1000
        .ColWidth(9) = 0
    End With
End Sub

'untuk cek validasi
Private Function funcCekValidasi() As Boolean
'    If  = "" Then
'        MsgBox "Pilihan Pelayanan Pasien Harus Diisi", vbCritical, "Validasi"
'        funcCekValidasi = False
'        txtNamaPelayanan.SetFocus
'        Exit Function
'    End If
    If fgPelayanan.TextMatrix(1, 0) = "" Then
        MsgBox "Pilihan Pelayanan Pasien Harus Diisi", vbCritical, "Validasi"
        funcCekValidasi = False
        txtNamaPelayanan.SetFocus
        Exit Function
    End If
    funcCekValidasi = True
End Function

'untuk enable/disable button reg
Private Sub subEnableButtonReg(blnStatus As Boolean)
    fraPDokter.Enabled = blnStatus
    fraPPelayanan.Enabled = blnStatus
'    fraButton.Enabled = blnStatus
    fgPelayanan.Enabled = blnStatus
    fgDOA.Enabled = blnStatus
    cmdSimpan.Enabled = blnStatus
End Sub

'untuk mengecek stok barang
'-----------------------**--Yang ditambah--**-----------------------
Private Function funcCekStokBarang(intBarang As Integer, strSatuanJml As String, intJml As Integer) As Boolean
    If strSatuanJml = "S" Then
        'paket layanan memakai satuan besar
        If (intJml * typBarang(intBarang).intJmlTerkecil) > _
        typBarang(intBarang).intJmlTempTotal Then
            MsgBox "Stok Barang '" & typBarang(intBarang).strNamaBarang & "' Tidak Cukup !", vbCritical, "Validasi"
            funcCekStokBarang = False
            Exit Function
        Else
            typBarang(intBarang).intJmlTempTotal = typBarang(intBarang).intJmlTempTotal - (intJml * typBarang(intBarang).intJmlTerkecil)
        End If
    Else
        If intJml > typBarang(intBarang).intJmlTempTotal Then
            MsgBox "Stok Barang '" & typBarang(intBarang).strNamaBarang & "' Tidak Cukup !", vbCritical, "Validasi"
            funcCekStokBarang = False
            Exit Function
        Else
            typBarang(intBarang).intJmlTempTotal = typBarang(intBarang).intJmlTempTotal - intJml
        End If
    End If
    funcCekStokBarang = True
End Function
'-----------------------**--Yang ditambah--**-----------------------
Private Sub txtNamaPerawat_Change()
On Error GoTo errLoad

    Call subLoadListPemeriksa("where v_DaftarPemeriksaPasien.[Nama Pemeriksa] LIKE '%" & txtNamaPerawat.Text & "%'")
    lvPemeriksa.Visible = True

Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtNamaPerawat_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvPemeriksa.Visible = True Then If lvPemeriksa.ListItems.Count > 0 Then lvPemeriksa.SetFocus
        Case vbKeyEscape
            lvPemeriksa.Visible = False
    End Select
End Sub

Private Sub txtNamaPerawat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lvPemeriksa.Visible = True Then
            lvPemeriksa.SetFocus
        Else
            optCito(1).SetFocus
        End If
    End If
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 6
        .Rows = 1
        
        .MergeCells = flexMergeFree
        
        .TextMatrix(0, 0) = "NoPendaftaran"
        .TextMatrix(0, 1) = "Kode Ruangan"
        .TextMatrix(0, 2) = "Tgl Pelayanan"
        .TextMatrix(0, 3) = "Kode Pelayanan"
        .TextMatrix(0, 4) = "IdPegawai"
        .TextMatrix(0, 5) = "IdUser"
    End With
End Sub

Private Sub subLoadListPemeriksa(Optional strKriteria As String)
Dim strKey As String
    
    strSQL = "SELECT     v_DaftarPemeriksaPasien.IdPegawai, v_DaftarPemeriksaPasien.[Nama Pemeriksa], v_DaftarPemeriksaPasien.JK, v_DaftarPemeriksaPasien.[Jenis Pemeriksa], v_DaftarPemeriksaPasien.StatusEnabled FROM         v_DaftarPemeriksaPasien INNER JOIN JenisPegawai ON v_DaftarPemeriksaPasien.[Jenis Pemeriksa] = JenisPegawai.JenisPegawai INNER JOIN SettingGlobal ON JenisPegawai.KdJenisPegawai = SettingGlobal.Value " & strKriteria & " and  (SettingGlobal.Prefix = 'KdJenisPegawaiParamedis') order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)
    
    With lvPemeriksa
        .ListItems.clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).Value
            .ListItems.add , strKey, rs(1).Value
            rs.MoveNext
        Next
    
        .Top = fraPDokter.Top + txtNamaPerawat.Top + txtNamaPerawat.Height
        .Left = fraPDokter.Left + txtNamaPerawat.Left
        .Height = 1815
        .ColumnHeaders.Item(1).Width = lvPemeriksa.Width - 500
        
        If subJmlTotal = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotal
                If .ListItems(i).key = subKdPemeriksa(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Function sp_Take_TarifBPT() As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdPelayananRS", adChar, adParamInput, 6, strKodePelayananRS)
        .Parameters.Append .CreateParameter("KdKelas", adChar, adParamInput, 2, mstrKdKelas)
        .Parameters.Append .CreateParameter("KdJenisTarif", adChar, adParamInput, 2, strKdJenisTarif)
        .Parameters.Append .CreateParameter("TarifCito", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)
        .Parameters.Append .CreateParameter("StatusCito", adChar, adParamInput, 1, IIf(optCito(0).Value = True, "Y", "T"))
        .Parameters.Append .CreateParameter("IdDokter", adChar, adParamInput, 10, IIf(chkDilayaniDokter.Value = vbChecked, mstrKdDokter, Null))
        .Parameters.Append .CreateParameter("IdDokter2", adChar, adParamInput, 10, Null)
        .Parameters.Append .CreateParameter("IdDokter3", adChar, adParamInput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifBPT"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifBPT = 0
            subcurTarifBiayaSatuan = 0
        Else
            sp_Take_TarifBPT = .Parameters("TarifCito").Value
            subcurTarifBiayaSatuan = .Parameters("TarifTotal").Value
            Call Add_HistoryLoginActivity("Take_TarifBPT")
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_Take_TarifOA(f_KdAsal As String, f_HargaSatuan As Currency) As Currency
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoPendaftaran", adChar, adParamInput, 10, mstrNoPen)
        .Parameters.Append .CreateParameter("KdAsal", adChar, adParamInput, 6, f_KdAsal)
        .Parameters.Append .CreateParameter("HargaSatuan", adCurrency, adParamInput, , CCur(f_HargaSatuan))
        .Parameters.Append .CreateParameter("TarifTotal", adCurrency, adParamOutput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "dbo.Take_TarifOA"
        .CommandType = adCmdStoredProc
        .Execute
        
        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Ada kesalahan dalam Pengambilan biaya tarif", vbExclamation, "Validasi"
            sp_Take_TarifOA = 0
        Else
            sp_Take_TarifOA = .Parameters("TarifTotal").Value
            Call Add_HistoryLoginActivity("Take_TarifOA")
        End If
    
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function


