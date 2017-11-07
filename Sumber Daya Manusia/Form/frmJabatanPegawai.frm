VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmJabatanPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Jenis Jabatan & Jabatan Pegawai"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "frmJabatanPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   8775
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   1320
      TabIndex        =   32
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   16
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   15
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   8055
      Width           =   1335
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jenis Jabatan"
      TabPicture(0)   =   "frmJabatanPegawai.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jabatan"
      TabPicture(1)   =   "frmJabatanPegawai.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6135
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   8055
         Begin VB.TextBox txtParameter 
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
            Left            =   4920
            MaxLength       =   30
            TabIndex        =   33
            Top             =   5720
            Width           =   2895
         End
         Begin VB.TextBox txtKdExt 
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
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   24
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkSts 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6120
            TabIndex        =   23
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNmExt 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   22
            Top             =   1320
            Width           =   5775
         End
         Begin VB.TextBox txtJenisJabatan 
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
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   5
            Top             =   600
            Width           =   4335
         End
         Begin VB.TextBox txtKdJenisJabatan 
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
            Height          =   315
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid dgJenisJabatan 
            Height          =   3840
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   6773
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cari Jenis Jabatan"
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
            Left            =   3120
            TabIndex        =   34
            Top             =   5760
            Width           =   1425
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
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
            Index           =   6
            Left            =   240
            TabIndex        =   26
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
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
            Index           =   7
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Jabatan"
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
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jenis"
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
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6135
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   8055
         Begin VB.TextBox txtParameterJabatan 
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
            Left            =   4800
            MaxLength       =   50
            TabIndex        =   35
            Top             =   5720
            Width           =   3015
         End
         Begin VB.TextBox txtKdExt1 
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
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   29
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CheckBox chkSts1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   6480
            TabIndex        =   28
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox txtNmExt1 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   27
            Top             =   2040
            Width           =   5775
         End
         Begin VB.TextBox txtNoUrut 
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
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   20
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtKdJabatan 
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
            Height          =   315
            Left            =   1920
            MaxLength       =   5
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtNamaJabatan 
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
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   8
            Top             =   600
            Width           =   5655
         End
         Begin MSDataListLib.DataCombo dcJenisJabatan 
            Height          =   315
            Left            =   1920
            TabIndex        =   9
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataGridLib.DataGrid dgJabatan 
            Height          =   3105
            Left            =   240
            TabIndex        =   19
            Top             =   2520
            Width           =   7560
            _ExtentX        =   13335
            _ExtentY        =   5477
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cari Nama Jabatan"
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
            Left            =   2880
            TabIndex        =   36
            Top             =   5760
            Width           =   1485
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
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
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   1680
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
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
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   2040
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "No. Urut"
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
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   705
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Jabatan"
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
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nama Jabatan"
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
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1140
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Kode Jabatan"
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
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1110
         End
      End
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
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmJabatanPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmJabatanPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmJabatanPegawai.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmJabatanPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDcSource()
    strSQL = "SELECT * FROM JenisJabatan order by JenisJabatan"
    Call msubDcSource(dcJenisJabatan, rs, strSQL)
End Sub

Sub sp_simpan()
    Select Case sstDataPenunjang.Tab
        Case 0 ' Jenis jabatan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisJabatan", adVarChar, adParamInput, 2, Trim(txtKdJenisJabatan))
                .Parameters.Append .CreateParameter("JenisJabatan", adVarChar, adParamInput, 30, Trim(txtJenisJabatan))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExt.Text)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExt.Text)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

                .ActiveConnection = dbConn
                .CommandText = "AUD_JenisJabatan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data ", vbExclamation, "Validasi"
                Else
                    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

        Case 1 ' jabatan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, Trim(txtKdJabatan))
                .Parameters.Append .CreateParameter("NamaJabatan", adVarChar, adParamInput, 50, Trim(txtNamaJabatan))
                .Parameters.Append .CreateParameter("KdJenisJabatan", adVarChar, adParamInput, 2, Trim(dcJenisJabatan.BoundText))
                .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, IIf(txtNoUrut.Text = "", Null, txtNoUrut.Text))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExt1.Text)
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExt1.Text)
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts1.Value)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")

                .ActiveConnection = dbConn
                .CommandText = "AUD_Jabatan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam penyimpanan data", vbExclamation, "Validasi"
                Else
                    MsgBox "Data berhasil disimpan ", vbInformation, "Informasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click
    End Select
End Sub

Private Sub cmdBatal_Click()
    Call subKosong
    Call subLoadGridSource
    Call subDcSource
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    '    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0
            If dgJenisJabatan.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakJenisJabatan.Show
        Case 1
            If dgJabatan.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakJabatan.Show
    End Select
hell:
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo hell
    Select Case sstDataPenunjang.Tab
        Case 0 'Jenis jabatan
            If Periksa("text", txtJenisJabatan, "Isi jenis jabatan!") = False Then Exit Sub
            If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            Set rs = Nothing
            strSQL = "delete JenisJabatan where KdJenisJabatan = '" & txtKdJenisJabatan & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing

        Case 1 'jabatan
            If Periksa("datacombo", dcJenisJabatan, "Isi jenis jabatan") = False Then Exit Sub
            If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
            Set rs = Nothing
            strSQL = "delete Jabatan where KdJabatan = '" & txtKdJabatan & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
    End Select
    Call cmdBatal_Click
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Exit Sub
hell:
    MsgBox "tidak bisa di hapus, data sudah di pakai  ", vbInformation, "Informasi"
    'Call msubPesanError
End Sub

Private Sub cmdSimpan_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'Jenis jabatan
            If Periksa("text", txtJenisJabatan, "Isi jenis jabatan!") = False Then Exit Sub
            Call sp_simpan

        Case 1 ' jabatan
            If dcJenisJabatan.Text <> "" Then
                If Periksa("datacombo", dcJenisJabatan, "Jenis Jabatan Tidak Terdaftar") = False Then Exit Sub
            End If
            
            If Periksa("datacombo", dcJenisJabatan, "Isi jenis jabatan") = False Then Exit Sub
            If Periksa("text", txtNamaJabatan, "Isi Nama jabatan") = False Then Exit Sub
            Call sp_simpan

    End Select
    Call cmdBatal_Click
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgJabatan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJabatan
    WheelHook.WheelHook dgJabatan
End Sub

Private Sub dgJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJabatan.Text = dgJabatan.Columns(0).Value
    txtNamaJabatan.Text = dgJabatan.Columns(1).Value
    If IsNull(dgJabatan.Columns(3)) Then dcJenisJabatan.BoundText = "" Else dcJenisJabatan.BoundText = dgJabatan.Columns(3)
    If IsNull(dgJabatan.Columns(4)) Then txtNoUrut.Text = "" Else txtNoUrut.Text = dgJabatan.Columns(4)
    txtKdExt1.Text = dgJabatan.Columns(5).Value
    txtNmExt1.Text = dgJabatan.Columns(6).Value
    chkSts1.Value = dgJabatan.Columns(7).Value
End Sub

Private Sub dcJenisJabatan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then txtNoUrut.SetFocus

On Error GoTo hell
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJenisJabatan.Text)) = 0 Then txtNoUrut.SetFocus: Exit Sub
        If dcJenisJabatan.MatchedWithList = True Then txtNoUrut.SetFocus: Exit Sub
        strSQL = "SELECT KdJenisJabatan,JenisJabatan FROM JenisJabatan WHERE (JenisJabatan LIKE '%" & dcJenisJabatan.Text & "%')"
        Call msubRecFO(rs, strSQL)
        If rs.EOF = True Then Exit Sub
        dcJenisJabatan.BoundText = rs(0).Value
        dcJenisJabatan.Text = rs(1).Value
    End If
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub dgJenisJabatan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisJabatan
    WheelHook.WheelHook dgJenisJabatan
End Sub

Private Sub dgJenisJabatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJenisJabatan.Text = dgJenisJabatan.Columns(0).Value
    txtJenisJabatan.Text = dgJenisJabatan.Columns(1).Value
    txtKdExt.Text = dgJenisJabatan.Columns(2).Value
    txtNmExt.Text = dgJenisJabatan.Columns(3).Value
    chkSts.Value = dgJenisJabatan.Columns(4).Value
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    sstDataPenunjang.Tab = 0
    Call cmdBatal_Click
End Sub

Sub subKosong()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 'Jenis jabatan
            txtKdJenisJabatan.Text = ""
            txtJenisJabatan.Text = ""
            txtJenisJabatan.SetFocus
            txtKdExt.Text = ""
            chkSts.Value = 1
            txtNmExt.Text = ""
        Case 1 'jabatan
            txtKdJabatan.Text = ""
            dcJenisJabatan.BoundText = ""
            txtNamaJabatan.Text = ""
            txtNamaJabatan.SetFocus
            txtNoUrut.Text = ""
            txtKdExt1.Text = ""
            chkSts1.Value = 1
            txtNmExt1.Text = ""
    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDcSource
    Call cmdBatal_Click
End Sub

Private Sub txtJenisJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Sub subLoadGridSource()
    On Error GoTo hell
    Select Case sstDataPenunjang.Tab
        Case 0 ' Jenis jabatan
            Set rs = Nothing
            strSQL = "select KdJenisJabatan AS [Kode Jenis], JenisJabatan AS [Jenis Jabatan], KodeExternal, NamaExternal, StatusEnabled from JenisJabatan where JenisJabatan LIKE '%" & txtParameter.Text & "%'"
            Call msubRecFO(rs, strSQL)
            Set dgJenisJabatan.DataSource = rs
            dgJenisJabatan.Columns(0).Width = 2200
            dgJenisJabatan.Columns(0).Alignment = vbCenter
            dgJenisJabatan.Columns(1).Width = 4800
            dgJenisJabatan.Columns(4).Width = 1250
        Case 1  'Jabatan
            Set rs = Nothing
            strSQL = "SELECT dbo.Jabatan.KdJabatan AS Kode, dbo.Jabatan.NamaJabatan AS [Nama Jabatan], dbo.JenisJabatan.JenisJabatan AS [Jenis Jabatan], dbo.Jabatan.KdJenisJabatan, dbo.Jabatan.NoUrut AS [No. Urut], " & _
            " dbo.Jabatan.KodeExternal,dbo.Jabatan.NamaExternal,dbo.Jabatan.StatusEnabled FROM dbo.Jabatan LEFT OUTER JOIN" & _
            " dbo.JenisJabatan ON dbo.Jabatan.KdJenisJabatan = dbo.JenisJabatan.KdJenisJabatan where dbo.Jabatan.NamaJabatan LIKE '%" & txtParameterJabatan.Text & "%' "
            Call msubRecFO(rs, strSQL)
            Set dgJabatan.DataSource = rs
            dgJabatan.Columns(0).Width = 1000
            dgJabatan.Columns(0).Alignment = vbCenter
            dgJabatan.Columns(1).Width = 3500
            dgJabatan.Columns(2).Width = 1200
            dgJabatan.Columns(3).Width = 0
            dgJabatan.Columns(4).Width = 1000
            dgJabatan.Columns(7).Width = 1250
    End Select
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt.SetFocus
End Sub

Private Sub txtKdExt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt1.SetFocus
End Sub

Private Sub txtNamaJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisJabatan.SetFocus
End Sub

Private Sub txtNmExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNoUrut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt1.SetFocus
End Sub

Private Sub txtParameter_Change()
    Call subLoadGridSource
    strCetak = " where JenisJabatan LIKE '%" & txtParameter.Text & "%'"
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtParameterJabatan_Change()
    Call subLoadGridSource
    strCetak = " where dbo.Jabatan.NamaJabatan LIKE '%" & txtParameterJabatan.Text & "%'"
End Sub

Private Sub txtParameterJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
