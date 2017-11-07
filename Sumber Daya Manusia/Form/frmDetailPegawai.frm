VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmDetailPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Detail Pegawai"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   Icon            =   "frmDetailPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10455
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
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
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   5880
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
      Left            =   6480
      TabIndex        =   14
      Top             =   5880
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
      Left            =   9120
      TabIndex        =   16
      Top             =   5880
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
      Left            =   5160
      TabIndex        =   13
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Frame tdbTinggiBadan 
      Caption         =   "Detail Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   10215
      Begin VB.TextBox txtBeratBadan 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtTinggiBadan 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtCiriKhas 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4320
         MaxLength       =   300
         TabIndex        =   12
         Top             =   1800
         Width           =   5535
      End
      Begin VB.TextBox txtCacatTubuh 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         MaxLength       =   200
         ScrollBars      =   1  'Horizontal
         TabIndex        =   11
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtHobby 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   6240
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtWarnaKulit 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtBentukMuka 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtJenisRambut 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   240
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Ciri Khas"
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
         Left            =   4320
         TabIndex        =   34
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Cacat Tubuh"
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
         TabIndex        =   33
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Hobby"
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
         Left            =   6240
         TabIndex        =   32
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Warna Kulit"
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
         Left            =   4320
         TabIndex        =   31
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Bentuk Muka"
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
         Left            =   2040
         TabIndex        =   30
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "kg"
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
         Left            =   3480
         TabIndex        =   29
         Top             =   600
         Width           =   195
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "cm"
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
         Left            =   1680
         TabIndex        =   28
         Top             =   600
         Width           =   240
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Berat Badan"
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
         Left            =   2040
         TabIndex        =   26
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Tinggi Badan"
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
         TabIndex        =   25
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Rambut"
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
         TabIndex        =   24
         Top             =   960
         Width           =   960
      End
   End
   Begin VB.Frame fraDataPegawai 
      Caption         =   "Data Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   10215
      Begin VB.TextBox txtIdPegawai 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNamaLengkap 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtJenisPegawai 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   7320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtJK 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID Pegawai"
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
         TabIndex        =   22
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "JK"
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
         TabIndex        =   20
         Top             =   360
         Width           =   165
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Pegawai"
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
         Left            =   5160
         TabIndex        =   19
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Left            =   7320
         TabIndex        =   18
         Top             =   360
         Width           =   585
      End
   End
   Begin MSDataGridLib.DataGrid dgDetailPegawai 
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   1720
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
      TabIndex        =   35
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
      Picture         =   "frmDetailPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDetailPegawai.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDetailPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmDetailPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String

Private Sub cmdBatal_Click()
    Call subClearData
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo errHapus
    If MsgBox("Hapus Detail Pegawai ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    strSQL = "DELETE FROM DetailPegawai WHERE IdPegawai='" & txtIdPegawai.Text & "'"
    dbConn.Execute strSQL

    Call subLoadDetailPegawai
    Call subClearData
    Exit Sub
errHapus:
    Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad
    If Periksa("text", txtIdPegawai, "Nama Pegawai harus diisi dengan benar") = False Then Exit Sub
    If Periksa("text", txtTinggiBadan, "Tinggi Badan Pegawai harus diisi") = False Then Exit Sub
    If Periksa("Text", txtBeratBadan, "Berat Badan Pegawai harus diisi") = False Then Exit Sub
    
'    If MsgBox("Simpan data Detail Pegawai..", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then
'        Call cmdBatal_Click
'        Exit Sub
'    End If
    
    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, txtIdPegawai.Text)
        .Parameters.Append .CreateParameter("Hobby", adVarChar, adParamInput, 100, txtHobby.Text)
        .Parameters.Append .CreateParameter("TinggiBadan", adVarChar, adParamInput, 20, txtTinggiBadan.Text)
        .Parameters.Append .CreateParameter("BeratBadan", adVarChar, adParamInput, 20, txtBeratBadan.Text)
        .Parameters.Append .CreateParameter("JenisRambut", adVarChar, adParamInput, 20, txtJenisRambut.Text)
        .Parameters.Append .CreateParameter("BentukMuka", adVarChar, adParamInput, 50, txtBentukMuka.Text)
        .Parameters.Append .CreateParameter("WarnaKulit", adVarChar, adParamInput, 50, txtWarnaKulit.Text)
        .Parameters.Append .CreateParameter("CiriCiriKhas", adVarChar, adParamInput, 300, txtCiriKhas.Text)
        .Parameters.Append .CreateParameter("CacatTubuh", adVarChar, adParamInput, 200, txtCacatTubuh.Text)

        .ActiveConnection = dbConn
        .CommandText = "AU_HRD_DetailPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Detail pegawai", vbCritical
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            Exit Sub
        Else
            MsgBox "Data berhasil disimpan...", vbInformation, "Informasi"
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
    Call subLoadDetailPegawai
    Call subClearData
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    frmDataPegawaiNew.Enabled = True
    Unload Me
End Sub

Private Sub dgDetailPegawai_Click()
    If dgDetailPegawai.ApproxCount = 0 Then Exit Sub
    cmdHapus.Enabled = True
    cmdSimpan.Enabled = True
    If dgDetailPegawai.Columns(1) = "" Then txtHobby.Text = "" Else txtHobby.Text = dgDetailPegawai.Columns(1).Value
    If dgDetailPegawai.Columns(2) = "" Then txtTinggiBadan.Text = "" Else txtTinggiBadan.Text = dgDetailPegawai.Columns(2).Value
    If dgDetailPegawai.Columns(3) = "" Then txtBeratBadan.Text = "" Else txtBeratBadan.Text = dgDetailPegawai.Columns(3).Value
    If dgDetailPegawai.Columns(4) = "" Then txtJenisRambut.Text = "" Else txtJenisRambut.Text = dgDetailPegawai.Columns(4).Value
    If dgDetailPegawai.Columns(5) = "" Then txtBentukMuka.Text = "" Else txtBentukMuka.Text = dgDetailPegawai.Columns(5).Value
    If dgDetailPegawai.Columns(6) = "" Then txtWarnaKulit.Text = "" Else txtWarnaKulit.Text = dgDetailPegawai.Columns(6).Value
    If dgDetailPegawai.Columns(7) = "" Then txtCiriKhas.Text = "" Else txtCiriKhas.Text = dgDetailPegawai.Columns(7).Value
    If dgDetailPegawai.Columns(8) = "" Then txtCacatTubuh.Text = "" Else txtCacatTubuh.Text = dgDetailPegawai.Columns(8).Value
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subClearData
    Call subLoadDetailPegawai
    cmdSimpan.Enabled = True
    Call subLoadDataPegawai
    txtJenisRambut.SetFocus
End Sub

Private Sub subLoadDetailPegawai()
    On Error GoTo errLoad
    strLSQL = "SELECT * FROM DetailPegawai WHERE IdPegawai='" & mstrIdPegawai & "'"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    With dgDetailPegawai
        Set .DataSource = rs
        .Columns(0).Width = 0           'IdPegawai

        .Columns(1).Width = 2000
        .Columns(2).Width = 2000
        .Columns(3).Width = 2000
        .Columns(4).Width = 2000
        .Columns(5).Width = 2000
        .Columns(6).Width = 2000
        .Columns(7).Width = 2000
        .Columns(8).Width = 2000

        .Columns(1).Caption = "Hobby"
        .Columns(2).Caption = "Tinggi Badan"
        .Columns(3).Caption = "Berat Badan"
        .Columns(4).Caption = "Jenis Rambut"
        .Columns(5).Caption = "Bentuk Muka"
        .Columns(6).Caption = "Warna Kulit"
        .Columns(7).Caption = "Ciri Khas"
        .Columns(8).Caption = "Cacat Tubuh"
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subClearData()
    txtBeratBadan.Text = ""
    txtJenisRambut.Text = ""
    txtBentukMuka.Text = ""
    txtWarnaKulit.Text = ""
    txtHobby.Text = ""
    txtCacatTubuh.Text = ""
    txtCiriKhas.Text = ""
    txtTinggiBadan.Text = ""
    txtBeratBadan.Text = ""
End Sub

Private Sub subLoadDataPegawai()
    On Error GoTo errLoad
    strLSQL = "SELECT * FROM v_S_Pegawai WHERE IdPegawai = '" & mstrIdPegawai & "'"
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = True Then
        MsgBox "Nama Pegawai tidak ada", vbCritical, "Validasi"
        Exit Sub
    Else
        txtIdPegawai.Text = rs.Fields("IdPegawai")
        txtJK.Text = rs.Fields("Sex")
        If IsNull(rs.Fields("JenisPegawai")) Then
            txtJenisPegawai.Text = ""
        Else
            txtJenisPegawai.Text = rs.Fields("JenisPegawai")
        End If
        If IsNull(rs.Fields("NamaJabatan")) Then
            txtJabatan.Text = ""
        Else
            txtJabatan.Text = rs.Fields("NamaJabatan")
        End If
        Call subLoadDetailPegawai
    End If
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDataPegawaiNew.Enabled = True
End Sub

Private Sub txtBeratBadan_Change()
    txtBeratBadan.MaxLength = 3
End Sub

Private Sub txtNamaLengkap_Change()
    Call subLoadDataPegawai
End Sub

Private Sub txtNamaLengkap_GotFocus()
    txtNamaLengkap.SelStart = 0
    txtNamaLengkap.SelLength = Len(txtNamaLengkap.Text)
End Sub

Private Sub txtNamaLengkap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call subLoadDataPegawai
        cmdSimpan.Enabled = True
    End If
End Sub

Private Sub txtTinggiBadan_Change()
    txtTinggiBadan.MaxLength = 3
End Sub

Private Sub txtTinggiBadan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBeratBadan.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtBeratBadan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisRambut.SetFocus
    If Not (KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtJenisRambut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtBentukMuka.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtJenisRambut_LostFocus()
    txtJenisRambut.Text = StrConv(txtJenisRambut.Text, vbProperCase)
End Sub

Private Sub txtBentukMuka_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtWarnaKulit.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtBentukMuka_LostFocus()
    txtBentukMuka.Text = StrConv(txtBentukMuka.Text, vbProperCase)

End Sub

Private Sub txtWarnaKulit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtHobby.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtWarnaKulit_LostFocus()
    txtWarnaKulit.Text = StrConv(txtWarnaKulit.Text, vbProperCase)
End Sub

Private Sub txtHobby_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtCacatTubuh.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtHobby_LostFocus()
    txtHobby.Text = StrConv(txtHobby.Text, vbProperCase)
End Sub

Private Sub txtCacatTubuh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then txtCiriKhas.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtCacatTubuh_LostFocus()
    txtCacatTubuh.Text = StrConv(txtCacatTubuh.Text, vbProperCase)
End Sub

Private Sub txtCiriKhas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToChar(KeyAscii)
End Sub

Private Sub txtCiriKhas_LostFocus()
    txtCiriKhas.Text = StrConv(txtCiriKhas.Text, vbProperCase)
End Sub
