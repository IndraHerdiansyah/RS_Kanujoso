VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmUsulanRealisasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Riwayat Usulan Pegawai"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUsulanRealisasi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   12615
   Begin MSDataGridLib.DataGrid dgPegawai 
      Height          =   3060
      Left            =   360
      TabIndex        =   19
      Top             =   4080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   5398
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   19
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
         Size            =   9.75
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
   Begin VB.TextBox txtNoRiwayat 
      Height          =   315
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtNamaFormPengirim 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   330
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame fraDataBarang 
      Caption         =   "Data Usulan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   20
      Top             =   3480
      Width           =   12375
      Begin MSDataListLib.DataCombo dcDKategory 
         Height          =   330
         Left            =   8160
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcAlasan 
         Height          =   330
         Left            =   6840
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcStatus 
         Height          =   330
         Left            =   5520
         TabIndex        =   34
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcPangkat 
         Height          =   330
         Left            =   4200
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtIsi 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Left            =   2760
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fgData 
         Height          =   3615
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   6376
         _Version        =   393216
         FixedCols       =   0
         BackColorSel    =   -2147483643
         FocusRect       =   2
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Surat Keputusan Usulan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   960
      Width           =   12375
      Begin VB.Frame Frame1 
         Caption         =   "Jenis Usulan"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   12135
         Begin VB.OptionButton option1 
            Caption         =   "Kenaikan Pangkat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Kenaikan Gaji"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2040
            TabIndex        =   8
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Pensiun"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3600
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "TAPERUM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4800
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton option7 
            Caption         =   "Pemberhentian TPHL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   10200
            TabIndex        =   13
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton option6 
            Caption         =   "Pengangkatan PNS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   8160
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.OptionButton option5 
            Caption         =   "Pengangkatan TPHL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6120
            TabIndex        =   11
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox txtKeterangan 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6120
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1320
         Width           =   6015
      End
      Begin VB.TextBox txtTTDSK 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3000
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox txtNoSK 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   360
         MaxLength       =   30
         TabIndex        =   0
         Top             =   600
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker dtpTglSK 
         Height          =   330
         Left            =   3960
         TabIndex        =   1
         Top             =   600
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
         CustomFormat    =   "dd MMM yyyy "
         Format          =   105119747
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSDataListLib.DataCombo dcPegawai 
         Height          =   330
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker dtpMulai 
         Height          =   330
         Left            =   6000
         TabIndex        =   2
         Top             =   600
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
         CustomFormat    =   "dd MMM yyyy "
         Format          =   105119747
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin MSComCtl2.DTPicker dtpAkhir 
         Height          =   330
         Left            =   8040
         TabIndex        =   3
         Top             =   600
         Width           =   2175
         _ExtentX        =   3836
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
         CheckBox        =   -1  'True
         CustomFormat    =   "dd MMM yyyy "
         Format          =   105119747
         UpDown          =   -1  'True
         CurrentDate     =   37760
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   5
         Left            =   6120
         TabIndex        =   31
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penanda Tangan SK Lain"
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
         Left            =   3000
         TabIndex        =   30
         Top             =   1080
         Width           =   2220
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Akhir Berlaku"
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
         Left            =   8040
         TabIndex        =   29
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Mulai Berlaku"
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
         Index           =   3
         Left            =   6000
         TabIndex        =   28
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl. Surat Keputusan"
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
         Left            =   3960
         TabIndex        =   24
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Surat Keputusan"
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
         Left            =   360
         TabIndex        =   23
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Penanda Tangan SK"
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
         Index           =   10
         Left            =   360
         TabIndex        =   22
         Top             =   1080
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   9240
      TabIndex        =   16
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   7560
      Width           =   1575
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   27
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
      Left            =   10800
      Picture         =   "frmUsulanRealisasi.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmUsulanRealisasi.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmUsulanRealisasi.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmUsulanRealisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tempStatusTampil As Boolean
Dim subbolSimpan As Boolean
Dim intBanyakData As String

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadDcSource
    Call subSetGrid
    txtNoSK.SetFocus
End Sub

Private Sub msg()
    MsgBox "Data Berhasil Disimpan", vbInformation, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    Dim i As Integer
    If Periksa("text", txtNoSK, "Silahkan isi Nomor Surat Keputusan ") = False Then Exit Sub
    If fgData.TextMatrix(1, 0) = "" Then MsgBox "Data Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub

    If sp_Riwayat("A") = False Then Exit Sub
    If sp_RiwayatSK() = False Then Exit Sub

    For i = 1 To fgData.Rows - 2
        If option1.Value = True Then
            If fgData.TextMatrix(i, 14) = "" Then MsgBox "Pangkat Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, Null)
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, fgData.TextMatrix(i, 14))
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With

        End If

        If Option2.Value = True Then
            If fgData.TextMatrix(i, 6) = "" Then MsgBox "Gaji Pokok Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, Null)
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Val(fgData.TextMatrix(i, 6)))
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With
        End If

        If Option3.Value = True Then
            If fgData.TextMatrix(i, 11) = "" Then MsgBox "Status Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            If fgData.TextMatrix(i, 12) = "" Then MsgBox "Alasan Status Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, Null)
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, fgData.TextMatrix(i, 11))
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , fgData.TextMatrix(i, 12))
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With
        End If

        If Option4.Value = True Then
            If fgData.TextMatrix(i, 8) = "" Then MsgBox "Total Pagu Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            If fgData.TextMatrix(i, 9) = "" Then MsgBox "Bank Penyalur Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, Null)
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Val(fgData.TextMatrix(i, 8)))
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, fgData.TextMatrix(i, 9))
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With
        End If

        If option5.Value = True Then
            If fgData.TextMatrix(i, 11) = "" Then MsgBox "Status Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, IIf(fgData.TextMatrix(i, 2) = "", Null, fgData.TextMatrix(i, 2)))
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, fgData.TextMatrix(i, 11))
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With
        End If

        If option6.Value = True Then
            If fgData.TextMatrix(i, 13) = "" Then MsgBox "Detail Kategory Pegawai Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            If fgData.TextMatrix(i, 6) = "" Then MsgBox "Gaji Pokok Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, IIf(fgData.TextMatrix(i, 2) = "", Null, fgData.TextMatrix(i, 2)))
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, fgData.TextMatrix(i, 13))
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Val(fgData.TextMatrix(i, 6)))
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With
        End If

        If option7.Value = True Then
            If fgData.TextMatrix(i, 11) = "" Then MsgBox "Status Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            If fgData.TextMatrix(i, 12) = "" Then MsgBox "Alasan Status Usulan harus diisi", vbExclamation, "Validasi": fgData.SetFocus: Exit Sub
            Set dbcmd = New ADODB.Command
            With adoComm
                .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
                .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, fgData.TextMatrix(i, 0))
                .Parameters.Append .CreateParameter("TugasPekerjaan", adVarChar, adParamInput, 150, IIf(fgData.TextMatrix(i, 2) = "", Null, fgData.TextMatrix(i, 2)))
                .Parameters.Append .CreateParameter("KdStatusUsulan", adChar, adParamInput, 2, fgData.TextMatrix(i, 11))
                .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , fgData.TextMatrix(i, 12))
                .Parameters.Append .CreateParameter("KdDKategoryPUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("GajiPokokUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("KdPangkatUsulan", adVarChar, adParamInput, 2, Null)
                .Parameters.Append .CreateParameter("NoRiwayatRealisasi", adChar, adParamInput, 10, Null)
                .Parameters.Append .CreateParameter("TotalPaguUsulan", adCurrency, adParamInput, , Null)
                .Parameters.Append .CreateParameter("BankPenyalur", adVarChar, adParamInput, 50, Null)
                .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 100, IIf(fgData.TextMatrix(i, 15) = "", Null, Trim(fgData.TextMatrix(i, 15))))

                .ActiveConnection = dbConn
                .CommandText = "Add_HRD_DetailRiwayatUsulan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("RETURN_VALUE").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
                    Call deleteADOCommandParameters(adoComm)
                    Set adoComm = Nothing
                End If
                Call msg
                Call deleteADOCommandParameters(adoComm)
                Set adoComm = Nothing
            End With
        End If
    Next i

    Call cmdBatal_Click
    subbolSimpan = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTambah_Click()
    fgData.Rows = fgData.Rows + 1
End Sub

Private Sub Cmdtambah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then fgData.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcPangkat_Change()
    fgData.TextMatrix(fgData.Row, 7) = dcPangkat.Text
    fgData.TextMatrix(fgData.Row, 14) = dcPangkat.BoundText
End Sub

Private Sub dcStatus_Change()
    fgData.TextMatrix(fgData.Row, 3) = dcStatus.Text
    fgData.TextMatrix(fgData.Row, 11) = dcStatus.BoundText
End Sub

Private Sub dcAlasan_Change()
    fgData.TextMatrix(fgData.Row, 4) = dcAlasan.Text
    fgData.TextMatrix(fgData.Row, 12) = dcAlasan.BoundText
End Sub

Private Sub dcDKategory_Change()
    fgData.TextMatrix(fgData.Row, 5) = dcDKategory.Text
    fgData.TextMatrix(fgData.Row, 13) = dcDKategory.BoundText
End Sub

Private Sub dcPangkat_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        Call dcPangkat_Change

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
        If fgData.TextMatrix(fgData.Rows - 1, 0) <> "" Then fgData.Rows = fgData.Rows + 1
        dcPangkat.Visible = False
        fgData.SetFocus
        fgData.Col = 15
    ElseIf KeyAscii = 27 Then

        dgPegawai.Visible = False
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcStatus_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        Call dcStatus_Change

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
        If fgData.TextMatrix(fgData.Rows - 1, 0) <> "" Then fgData.Rows = fgData.Rows + 1
        dcStatus.Visible = False
        fgData.SetFocus
        If option5.Value = True Then
            fgData.Col = 15
        Else
            fgData.Col = 4
        End If
    ElseIf KeyAscii = 27 Then

        dgPegawai.Visible = False
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcAlasan_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        Call dcAlasan_Change

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
        If fgData.TextMatrix(fgData.Rows - 1, 0) <> "" Then fgData.Rows = fgData.Rows + 1
        dcAlasan.Visible = False
        fgData.SetFocus
        fgData.Col = 15
    ElseIf KeyAscii = 27 Then

        dgPegawai.Visible = False
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcDKategory_KeyPress(KeyAscii As Integer)
    On Error GoTo hell
    If KeyAscii = 13 Then
        Call dcDKategory_Change

        If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
            fgData.SetFocus
            SendKeys "{DOWN}"
            Exit Sub
        End If
        fgData.SetFocus
        If fgData.TextMatrix(fgData.Rows - 1, 0) <> "" Then fgData.Rows = fgData.Rows + 1
        dcDKategory.Visible = False
        fgData.SetFocus
        fgData.Col = 6
    ElseIf KeyAscii = 27 Then

        dgPegawai.Visible = False
        fgData.SetFocus
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dcPangkat_LostFocus()
    dcPangkat.Visible = False
End Sub

Private Sub dcstatus_LostFocus()
    dcStatus.Visible = False
End Sub

Private Sub dcAlasan_LostFocus()
    dcAlasan.Visible = False
End Sub

Private Sub dcDKategory_LostFocus()
    dcDKategory.Visible = False
End Sub

Private Sub dcPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTTDSK.SetFocus
End Sub

Private Sub dgPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPegawai
    WheelHook.WheelHook dgPegawai
End Sub

Private Sub dgPegawai_DblClick()
    On Error GoTo errLoad
    Dim i As Integer

    tempStatusTampil = True
    With fgData
        dgPegawai.Visible = False
        .TextMatrix(.Row, 0) = dgPegawai.Columns(0)
        .TextMatrix(.Row, 1) = dgPegawai.Columns(1)

    End With
    tempStatusTampil = False
    fgData.SetFocus
    If option1.Value = True Then
        fgData.Col = 7
    ElseIf Option2.Value = True Then
        fgData.Col = 6
    ElseIf Option3.Value = True Then
        fgData.Col = 3
    ElseIf Option4.Value = True Then
        fgData.Col = 8
    ElseIf option5.Value = True Then
        fgData.Col = 2
    ElseIf option6.Value = True Then
        fgData.Col = 5
    ElseIf option7.Value = True Then
        fgData.Col = 2
    End If
    Exit Sub
errLoad:
End Sub

Private Sub dgPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then dgPegawai.Visible = False: fgData.SetFocus
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call dgPegawai_DblClick
End Sub

Private Sub dtpAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dcPegawai.SetFocus
End Sub

Private Sub dtpMulai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub dtpTglSK_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then dtpMulai.SetFocus
End Sub

Private Sub fgData_DblClick()
    Call fgData_KeyDown(13, 0)
End Sub

Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
        Case 13
            If fgData.Col = fgData.Cols - 1 Then
                If fgData.TextMatrix(fgData.Row, 1) <> "" Then
                    If fgData.TextMatrix(fgData.Rows - 1, 1) <> "" Then fgData.Rows = fgData.Rows + 1
                    fgData.Row = fgData.Rows - 1
                    fgData.Col = 1
                Else
                    fgData.Col = 1
                End If
            Else
                For i = 0 To fgData.Cols - 2
                    If fgData.Col = fgData.Cols - 1 Then Exit For
                    fgData.Col = fgData.Col + 1
                    If fgData.ColWidth(fgData.Col) > 0 Then Exit For
                Next i
            End If
            fgData.SetFocus

        Case 27
            dgPegawai.Visible = False

        Case vbKeyDelete
            With fgData
                If .Row = .Rows Then Exit Sub
                If .Row = 0 Then Exit Sub

                If .Rows = 2 Then
                    For i = 0 To .Cols - 1
                        .TextMatrix(1, i) = ""
                    Next i
                    Exit Sub
                Else
                    .RemoveItem .Row
                End If
            End With

    End Select
End Sub

Private Sub fgData_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad

    txtIsi.Text = ""
    If Not (KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii >= vbKeyA And KeyAscii <= vbKeyZ Or KeyAscii = 32 Or KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = Asc(".")) Then
        KeyAscii = 0
        Exit Sub
    End If

    If option1.Value = True Then
        Select Case fgData.Col

            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)

            Case 7 'usulan pangkat
                fgData.Col = 7
                Call subLoadDataCombo(dcPangkat)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    ElseIf Option2.Value = True Then
        Select Case fgData.Col
            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 6 'Gaji Pokok Usulan
                txtIsi.MaxLength = 8
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    ElseIf Option3.Value = True Then
        Select Case fgData.Col
            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)

            Case 3 'status usulan
                fgData.Col = 3
                Call subLoadDataCombo(dcStatus)
            Case 4 'alasan status usulan
                fgData.Col = 4
                Call subLoadDataCombo(dcAlasan)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    ElseIf Option4.Value = True Then
        Select Case fgData.Col
            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)

            Case 8 'pagu usulan
                txtIsi.MaxLength = 8
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 9 'bank penyalur
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    ElseIf option5.Value = True Then
        Select Case fgData.Col
            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)

            Case 2 'tugas pekerjaan
                txtIsi.MaxLength = 150
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 3 'status usulan
                fgData.Col = 3
                Call subLoadDataCombo(dcStatus)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    ElseIf option6.Value = True Then
        Select Case fgData.Col
            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 5 'Detail Kategory Pegawai usulan
                fgData.Col = 5
                Call subLoadDataCombo(dcDKategory)
            Case 6 'Gaji Pokok Usulan
                txtIsi.MaxLength = 8
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    ElseIf option7.Value = True Then
        Select Case fgData.Col
            Case 1 'Nama Pegawai
                txtIsi.MaxLength = 50
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 2 'tugas pekerjaan
                txtIsi.MaxLength = 150
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
            Case 3 'status usulan
                fgData.Col = 3
                Call subLoadDataCombo(dcStatus)
            Case 4 'alasan status usulan
                fgData.Col = 4
                Call subLoadDataCombo(dcAlasan)
            Case 15 'Keterangan
                txtIsi.MaxLength = 100
                Call subLoadText
                txtIsi.Text = Chr(KeyAscii)
                txtIsi.SelStart = Len(txtIsi.Text)
        End Select
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)

    dtpTglSK.Value = Now
    dtpMulai.Value = Now
    dtpAkhir.Value = Now
    option1.Value = True

    Call subKosong
    Call subSetGrid
    Call subLoadDcSource
    dcPangkat.BoundText = ""
    dcStatus.BoundText = ""
    dcAlasan.BoundText = ""
    dcDKategory.BoundText = ""
    dgPegawai.Visible = False
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Option1_Click()
    Call subSetGrid
End Sub

Private Sub Option2_Click()
    Call subSetGrid
End Sub

Private Sub Option3_Click()
    Call subSetGrid
End Sub

Private Sub Option4_Click()
    Call subSetGrid
End Sub

Private Sub Option5_Click()
    Call subSetGrid
End Sub

Private Sub Option6_Click()
    Call subSetGrid
End Sub

Private Sub Option7_Click()
    Call subSetGrid
End Sub

Private Sub txtIsi_Change()
    On Error GoTo errLoad
    Dim i As Integer
    Select Case fgData.Col

        Case 1
            If option1.Value = True Then
                If tempStatusTampil = True Then Exit Sub
                strSQL = "select DISTINCT TOP 100 IdPegawai, NamaLengkap AS Nama, NIP, NamaJabatan AS Jabatan" & _
                " FROM V_LoadDataPegawaiForUsulan " & _
                " where NamaLengkap like '" & txtIsi.Text & "%' and KdKategoryPegawai IN('1','2') ORDER BY NamaLengkap "
                Call msubRecFO(dbRst, strSQL)

                Set dgPegawai.DataSource = dbRst
                With dgPegawai
                    .Columns(0).Width = 0
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 2200
                    .Columns(3).Width = 2500

                    .Left = 360
                    .Top = 4400
                    .Visible = True
                    For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                    Next i
                    If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                    End If
                End With
            ElseIf Option3.Value = True Then
                If tempStatusTampil = True Then Exit Sub
                strSQL = "select DISTINCT TOP 100 IdPegawai, NamaLengkap AS Nama, NIP, NamaJabatan AS Jabatan" & _
                " FROM V_LoadDataPegawaiForUsulan " & _
                " where NamaLengkap like '" & txtIsi.Text & "%' and KdKategoryPegawai IN('2') ORDER BY NamaLengkap "
                Call msubRecFO(dbRst, strSQL)

                Set dgPegawai.DataSource = dbRst
                With dgPegawai
                    .Columns(0).Width = 0
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 2200
                    .Columns(3).Width = 2500

                    .Left = 360
                    .Top = 4400
                    .Visible = True
                    For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                    Next i
                    If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                    End If
                End With
            ElseIf option5.Value = True Then
                If tempStatusTampil = True Then Exit Sub
                strSQL = "select DISTINCT TOP 100 IdPegawai, NamaLengkap AS Nama, NIP, NamaJabatan AS Jabatan" & _
                " FROM V_LoadDataPegawaiForUsulan " & _
                " where NamaLengkap like '" & txtIsi.Text & "%' and KdKategoryPegawai IN('3') ORDER BY NamaLengkap "
                Call msubRecFO(dbRst, strSQL)

                Set dgPegawai.DataSource = dbRst
                With dgPegawai
                    .Columns(0).Width = 0
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 2200
                    .Columns(3).Width = 2500

                    .Left = 360
                    .Top = 4400
                    .Visible = True
                    For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                    Next i
                    If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                    End If
                End With
            ElseIf option6.Value = True Then
                If tempStatusTampil = True Then Exit Sub
                strSQL = "select DISTINCT TOP 100 IdPegawai, NamaLengkap AS Nama, NIP, NamaJabatan AS Jabatan" & _
                " FROM V_LoadDataPegawaiForUsulan " & _
                " where NamaLengkap like '" & txtIsi.Text & "%' and KdKategoryPegawai IN('1','2') ORDER BY NamaLengkap "
                Call msubRecFO(dbRst, strSQL)

                Set dgPegawai.DataSource = dbRst
                With dgPegawai
                    .Columns(0).Width = 0
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 2200
                    .Columns(3).Width = 2500

                    .Left = 360
                    .Top = 4400
                    .Visible = True
                    For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                    Next i
                    If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                    End If
                End With
            ElseIf option7.Value = True Then
                If tempStatusTampil = True Then Exit Sub
                strSQL = "select DISTINCT TOP 100 IdPegawai, NamaLengkap AS Nama, NIP, NamaJabatan AS Jabatan" & _
                " FROM V_LoadDataPegawaiForUsulan " & _
                " where NamaLengkap like '" & txtIsi.Text & "%' and KdKategoryPegawai IN('3') ORDER BY NamaLengkap "
                Call msubRecFO(dbRst, strSQL)

                Set dgPegawai.DataSource = dbRst
                With dgPegawai
                    .Columns(0).Width = 0
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 2200
                    .Columns(3).Width = 2500

                    .Left = 360
                    .Top = 4400
                    .Visible = True
                    For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                    Next i
                    If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                    End If
                End With
            Else
                If tempStatusTampil = True Then Exit Sub
                strSQL = "select DISTINCT TOP 100 IdPegawai, NamaLengkap AS Nama, NIP, NamaJabatan AS Jabatan" & _
                " FROM V_LoadDataPegawaiForUsulan " & _
                " where NamaLengkap like '" & txtIsi.Text & "%' and KdKategoryPegawai IN('1','2') ORDER BY NamaLengkap "
                Call msubRecFO(dbRst, strSQL)

                Set dgPegawai.DataSource = dbRst
                With dgPegawai
                    .Columns(0).Width = 0
                    .Columns(1).Width = 2000
                    .Columns(2).Width = 2200
                    .Columns(3).Width = 2500

                    .Left = 360
                    .Top = 4400
                    .Visible = True
                    For i = 1 To fgData.Row - 1
                        .Top = .Top + fgData.RowHeight(i)
                    Next i
                    If fgData.TopRow > 1 Then
                        .Top = .Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
                    End If
                End With
            End If
        Case Else
            dgPegawai.Visible = False
            Exit Sub
    End Select

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If dgPegawai.Visible = True Then dgPegawai.SetFocus
End Sub

Private Sub txtIsi_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case fgData.Col

            Case 1
                If dgPegawai.Visible = True Then
                    dgPegawai.SetFocus
                    Exit Sub
                Else
                    fgData.SetFocus
                    fgData.Col = 2
                End If

            Case 2
                If dgPegawai.Visible = True Then
                    dgPegawai.SetFocus
                    Exit Sub
                Else
                    fgData.SetFocus
                    fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
                    fgData.Col = 3
                End If

            Case 6
                If Val(txtIsi.Text) = 0 Then
                    txtIsi.Text = 0
                End If
                With fgData
                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    .SetFocus
                    .Col = 15
                End With

            Case 8
                If Val(txtIsi.Text) = 0 Then
                    txtIsi.Text = 0
                End If
                With fgData
                    .TextMatrix(.Row, .Col) = txtIsi.Text
                    .SetFocus
                    .Col = 9
                End With
            Case 9
                If dgPegawai.Visible = True Then
                    dgPegawai.SetFocus
                    Exit Sub
                Else
                    fgData.SetFocus
                    fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
                    fgData.Col = 15
                End If
            Case 15
                If dgPegawai.Visible = True Then
                    dgPegawai.SetFocus
                    Exit Sub
                Else
                    fgData.SetFocus
                    fgData.TextMatrix(fgData.Row, fgData.Col) = txtIsi.Text
                    fgData.Col = 15
                End If
                If fgData.RowPos(fgData.Row) >= fgData.Height - 360 Then
                    fgData.SetFocus
                    SendKeys "{DOWN}"
                    Exit Sub
                End If
                fgData.SetFocus
                If fgData.TextMatrix(fgData.Rows - 1, 0) <> "" Then fgData.Rows = fgData.Rows + 1
                fgData.SetFocus
                SendKeys "{DOWN}"
                fgData.Col = 1
        End Select

        txtIsi.Visible = False

    ElseIf KeyAscii = 27 Then
        txtIsi.Visible = False
        dgPegawai.Visible = False
        fgData.SetFocus
    End If

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub txtIsi_LostFocus()
    txtIsi.Visible = False
End Sub

Private Sub subKosong()
    txtNoSK.Text = ""
    dtpTglSK.Value = Now
    dtpMulai.Value = Now
    dtpAkhir.Value = Now
    dcPegawai.BoundText = ""
    dgPegawai.Visible = False
    dcPangkat.Visible = False
    dcStatus.Visible = False
    dcAlasan.Visible = False
    dcDKategory.Visible = False
    txtTTDSK.Text = ""
    txtKeterangan.Text = ""
End Sub

Private Sub subSetGrid()
    On Error GoTo errLoad
    With fgData
        .clear
        .Rows = 2
        .Cols = 16

        .RowHeight(0) = 400
        .TextMatrix(0, 0) = "IdPegawai"
        .TextMatrix(0, 1) = "Nama"
        .TextMatrix(0, 2) = "Tugas Pekerjaan"
        .TextMatrix(0, 3) = "Status Usulan"
        .TextMatrix(0, 4) = "Alasan Status"
        .TextMatrix(0, 5) = "Detail Kategori Usulan"
        .TextMatrix(0, 6) = "Gaji Pokok Usulan"
        .TextMatrix(0, 7) = "Pangkat Usulan"
        .TextMatrix(0, 8) = "Total Pagu Usulan"
        .TextMatrix(0, 9) = "Bank Penyalur"
        .TextMatrix(0, 10) = "NoRiwayatRealisasi"
        .TextMatrix(0, 11) = "KdStatusUsulan"
        .TextMatrix(0, 12) = "KdAlasanStatus"
        .TextMatrix(0, 13) = "KdDKategoryPUsulan"
        .TextMatrix(0, 14) = "KdPangkatUsulan"
        .TextMatrix(0, 15) = "Keterangan"

        If option1.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 3000
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 5000
        ElseIf Option2.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 0
            .ColWidth(6) = 3000
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 5000
        ElseIf Option3.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 0
            .ColWidth(3) = 2000
            .ColWidth(4) = 2000
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 5000
        ElseIf Option4.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 2000
            .ColWidth(9) = 2000
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 5000
        ElseIf option5.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 3000
            .ColWidth(3) = 2000
            .ColWidth(4) = 0
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 4000
        ElseIf option6.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 0
            .ColWidth(3) = 0
            .ColWidth(4) = 0
            .ColWidth(5) = 2000
            .ColWidth(6) = 2000
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 5000
        ElseIf option7.Value = True Then
            .ColWidth(0) = 0
            .ColWidth(1) = 3000
            .ColWidth(2) = 3000
            .ColWidth(3) = 1500
            .ColWidth(4) = 1500
            .ColWidth(5) = 0
            .ColWidth(6) = 0
            .ColWidth(7) = 0
            .ColWidth(8) = 0
            .ColWidth(9) = 0
            .ColWidth(10) = 0
            .ColWidth(11) = 0
            .ColWidth(12) = 0
            .ColWidth(13) = 0
            .ColWidth(14) = 0
            .ColWidth(15) = 3000
        End If

    End With

    Exit Sub
errLoad:
End Sub

Private Sub subLoadDcSource()
    On Error GoTo hell
    Call msubDcSource(dcPegawai, rs, "SELECT IdPegawai, NamaLengkap FROM DataPegawai ORDER BY NamaLengkap")
    Call msubDcSource(dcPangkat, rs, "SELECT KdPangkat, NamaPangkat FROM Pangkat ORDER BY NamaPangkat")
    If rs.EOF = False Then dcPangkat.BoundText = rs(0).Value
    Call msubDcSource(dcStatus, rs, "SELECT KdStatus, Status FROM StatusPegawai ORDER BY Status")
    If rs.EOF = False Then dcStatus.BoundText = rs(0).Value
    Call msubDcSource(dcAlasan, rs, "SELECT KdAlasanStatus, AlasanStatus FROM AlasanStatusPegawai ORDER BY AlasanStatus")
    If rs.EOF = False Then dcAlasan.BoundText = rs(0).Value
    Call msubDcSource(dcDKategory, rs, "SELECT KdDetailKategoryPegawai, DetailKategoryPegawai FROM DetailKategoryPegawai ORDER BY DetailKategoryPegawai")
    If rs.EOF = False Then dcDKategory.BoundText = rs(0).Value
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub subLoadText()
    Dim i As Integer
    txtIsi.Left = fgData.Left
    For i = 0 To fgData.Col - 1
        txtIsi.Left = txtIsi.Left + fgData.ColWidth(i)
    Next i
    txtIsi.Visible = True
    txtIsi.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        txtIsi.Top = txtIsi.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        txtIsi.Top = txtIsi.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    txtIsi.Width = fgData.ColWidth(fgData.Col)
    txtIsi.Height = fgData.RowHeight(fgData.Row)

    txtIsi.Visible = True
    txtIsi.SelStart = Len(txtIsi.Text)
    txtIsi.SetFocus
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
            MsgBox "Ada kesalahan dalam penyimpanan data nomor riwayat", vbCritical, "Validasi"
            sp_Riwayat = False
        Else
            If Not IsNull(.Parameters("Status").Value) Then txtNoRiwayat.Text = .Parameters("OutputNoRiwayat").Value
        End If

        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Function sp_RiwayatSK() As Boolean
    On Error GoTo hell
    sp_RiwayatSK = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("NoRiwayat", adChar, adParamInput, 10, Trim(txtNoRiwayat.Text))
        .Parameters.Append .CreateParameter("NoSK", adVarChar, adParamInput, 30, Trim(txtNoSK.Text))
        .Parameters.Append .CreateParameter("TglSK", adDate, adParamInput, , Format(dtpTglSK.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("TandaTanganSK", adVarChar, adParamInput, 50, IIf(txtTTDSK.Text = "", Null, Trim(txtTTDSK.Text)))
        .Parameters.Append .CreateParameter("IdTandaTanganSK", adChar, adParamInput, 10, IIf(dcPegawai.Text = "", Null, dcPegawai.BoundText))
        .Parameters.Append .CreateParameter("TglMulaiBerlakuSK", adDate, adParamInput, , Format(dtpMulai.Value, "yyyy/MM/dd"))
        If IsNull(dtpAkhir.Value) Then
            .Parameters.Append .CreateParameter("TglAkhirBerlakuSK", adDate, adParamInput, , Null)
        Else
            .Parameters.Append .CreateParameter("TglAkhirBerlakuSK", adDate, adParamInput, , Format(dtpAkhir.Value, "yyyy/MM/dd"))
        End If
        .Parameters.Append .CreateParameter("KeteranganLainnya", adVarChar, adParamInput, 150, IIf(txtKeterangan.Text = "", Null, Trim(txtKeterangan.Text)))
        .Parameters.Append .CreateParameter("TglDiSetujui", adDate, adParamInput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_RiwayatSuratKeputusan"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical, "Validasi"
            Call deleteADOCommandParameters(adoComm)
            Set adoComm = Nothing
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
    Exit Function
hell:
    Call msubPesanError
End Function

Private Sub subLoadDataCombo(s_DcName As Object)
    Dim i As Integer
    s_DcName.Left = fgData.Left
    For i = 0 To fgData.Col - 1
        s_DcName.Left = s_DcName.Left + fgData.ColWidth(i)
    Next i
    s_DcName.Visible = True
    s_DcName.Top = fgData.Top - 7

    For i = 0 To fgData.Row - 1
        s_DcName.Top = s_DcName.Top + fgData.RowHeight(i)
    Next i

    If fgData.TopRow > 1 Then
        s_DcName.Top = s_DcName.Top - ((fgData.TopRow - 1) * fgData.RowHeight(1))
    End If

    s_DcName.Width = fgData.ColWidth(fgData.Col)
    s_DcName.Height = fgData.RowHeight(fgData.Row)

    s_DcName.Visible = True
    s_DcName.SetFocus
End Sub

Private Sub txtKeterangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNoSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpTglSK.SetFocus
End Sub

Private Sub txtTTDSK_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKeterangan.SetFocus
End Sub

Public Function subLoadDataUsulan() As Boolean
    On Error GoTo errLoad
    Dim i As Integer
    Dim substrNoRiwayat As String
    If txtNamaFormPengirim.Text = "frmDaftarUsulanPegawaiMassal" Then
        If frmDaftarUsulanPegawaiMassal.option1.Value = True Then
            option1.Value = True
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            option5.Enabled = False
            option6.Enabled = False
            option7.Enabled = False
        ElseIf frmDaftarUsulanPegawaiMassal.Option2.Value = True Then
            Option2.Value = True
            option1.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            option5.Enabled = False
            option6.Enabled = False
            option7.Enabled = False
        ElseIf frmDaftarUsulanPegawaiMassal.Option3.Value = True Then
            Option3.Value = True
            Option2.Enabled = False
            option1.Enabled = False
            Option4.Enabled = False
            option5.Enabled = False
            option6.Enabled = False
            option7.Enabled = False
        ElseIf frmDaftarUsulanPegawaiMassal.option5.Value = True Then
            option5.Value = True
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            option1.Enabled = False
            option6.Enabled = False
            option7.Enabled = False
        ElseIf frmDaftarUsulanPegawaiMassal.option6.Value = True Then
            option6.Value = True
            Option2.Enabled = False
            Option3.Enabled = False
            Option4.Enabled = False
            option5.Enabled = False
            option1.Enabled = False
            option7.Enabled = False
        End If
    End If

    dgPegawai.Visible = False
    dcPangkat.Visible = False
    dcStatus.Visible = False
    dcAlasan.Visible = False
    dcDKategory.Visible = False
    Call subSetGrid

    strSQL = "SELECT * FROM V_DetailRiwayatUsulanPegawai WHERE [No.Riwayat] = '" & txtNoRiwayat.Text & "' and [No.Riwayat Realisasi] is null "
    Call msubRecFO(rs, strSQL)

    If rs.EOF = True Then
        txtNoSK.Text = ""
        dtpTglSK.Value = Now
        dtpMulai.Value = Now
        dtpAkhir.Value = Now
        dcPegawai.BoundText = ""
        txtTTDSK.Text = ""
        txtKeterangan.Text = ""
        substrNoRiwayat = ""
        subLoadDataUsulan = False
        Exit Function
    End If

    substrNoRiwayat = txtNoRiwayat.Text
    subLoadDataUsulan = True
    txtNoSK.Text = rs("No.SK Usulan").Value
    dtpTglSK.Value = rs("Tgl.SK Usulan").Value
    dtpMulai.Value = rs("TglMulaiBerlakuSK").Value
    If IsNull(rs("TglAkhirBerlakuSK")) Then dtpAkhir.Value = Null Else dtpAkhir.Value = rs("TglAkhirBerlakuSK").Value
    If IsNull(rs("IdTandaTanganSK")) Then dcPegawai.BoundText = "" Else dcPegawai.BoundText = rs("IdTandaTanganSK").Value
    If IsNull(rs("Tanda Tangan SK 1")) Then txtTTDSK.Text = "" Else txtTTDSK.Text = rs("Tanda Tangan SK 1").Value
    If IsNull(rs("Keterangan")) Then txtKeterangan.Text = "" Else txtKeterangan.Text = rs("Keterangan").Value
    With fgData
        For i = 1 To rs.RecordCount
            .TextMatrix(i, 0) = rs(2).Value
            .TextMatrix(i, 1) = rs(3).Value
            If IsNull(rs(8)) Then .TextMatrix(i, 2) = "" Else .TextMatrix(i, 2) = rs(8).Value
            If IsNull(rs(9)) Then .TextMatrix(i, 3) = "" Else .TextMatrix(i, 3) = rs(9).Value
            If IsNull(rs(10)) Then .TextMatrix(i, 4) = "" Else .TextMatrix(i, 4) = rs(10).Value
            If IsNull(rs(11)) Then .TextMatrix(i, 5) = "" Else .TextMatrix(i, 5) = rs(11).Value
            If IsNull(rs(15)) Then .TextMatrix(i, 6) = "" Else .TextMatrix(i, 6) = rs(15).Value
            If IsNull(rs(12)) Then .TextMatrix(i, 7) = "" Else .TextMatrix(i, 7) = rs(12).Value
            If IsNull(rs(17)) Then .TextMatrix(i, 8) = "" Else .TextMatrix(i, 8) = rs(17).Value
            If IsNull(rs(18)) Then .TextMatrix(i, 9) = "" Else .TextMatrix(i, 9) = rs(18).Value
            If IsNull(rs(19)) Then .TextMatrix(i, 10) = "" Else .TextMatrix(i, 10) = rs(19).Value
            If IsNull(rs(20)) Then .TextMatrix(i, 11) = "" Else .TextMatrix(i, 11) = rs(20).Value
            If IsNull(rs(21)) Then .TextMatrix(i, 12) = "" Else .TextMatrix(i, 12) = rs(21).Value
            If IsNull(rs(22)) Then .TextMatrix(i, 13) = "" Else .TextMatrix(i, 13) = rs(22).Value
            If IsNull(rs(23)) Then .TextMatrix(i, 14) = "" Else .TextMatrix(i, 14) = rs(23).Value
            If IsNull(rs(16)) Then .TextMatrix(i, 15) = "" Else .TextMatrix(i, 15) = rs(16).Value

            rs.MoveNext
            .Rows = .Rows + 1
        Next i
        .Row = 1
    End With

    Exit Function
errLoad:
    Call msubPesanError
End Function

