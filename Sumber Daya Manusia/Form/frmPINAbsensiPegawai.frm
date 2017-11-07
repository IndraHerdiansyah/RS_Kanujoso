VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPINAbsensiPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - PIN Absensi Pegawai"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPINAbsensiPegawai.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   9375
   Begin VB.Frame Frame3 
      Caption         =   "Transfer && Cek FingerPrint"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4320
      TabIndex        =   20
      Top             =   2160
      Width           =   4935
      Begin VB.CheckBox chkHapusDB 
         Alignment       =   1  'Right Justify
         Caption         =   "Hapus PIN di Data Base"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton cmdHapusPinFRS 
         Caption         =   "&Hapus PIN"
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
         Left            =   3480
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdCek 
         Caption         =   "&Cek PIN"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtCek 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Download Finger Print Data"
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   4695
         Begin VB.CommandButton cmdTransfer 
            Caption         =   "&Download"
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
            Left            =   3360
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtTujuan 
            Alignment       =   2  'Center
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
            Height          =   375
            Left            =   2160
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FRS Tujuan:"
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
            Index           =   7
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   885
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat FRS400:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
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
      TabIndex        =   46
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
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
      Left            =   5760
      TabIndex        =   35
      Top             =   8280
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3855
      Left            =   120
      TabIndex        =   29
      Top             =   4320
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Data Base"
      TabPicture(0)   =   "frmPINAbsensiPegawai.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dgPIN"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtParameter"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtBuatPIN"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdBuatPIN"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSimpanPIN"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tmrBuatPIN"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkTidakPunyaPIN"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "FRS"
      TabPicture(1)   =   "frmPINAbsensiPegawai.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ListView1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkPilihSemua"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtCari"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      Begin VB.CheckBox chkTidakPunyaPIN 
         Caption         =   "Tidak Memiliki PIN"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   2295
      End
      Begin VB.Timer tmrBuatPIN 
         Enabled         =   0   'False
         Interval        =   5
         Left            =   5520
         Top             =   3360
      End
      Begin VB.TextBox txtCari 
         Height          =   315
         Left            =   -68040
         TabIndex        =   48
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkPilihSemua 
         Caption         =   "Pilih Semua"
         Height          =   210
         Left            =   -74880
         TabIndex        =   47
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton cmdSimpanPIN 
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
         Left            =   8160
         TabIndex        =   45
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdBuatPIN 
         Caption         =   "B&uat PIN"
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
         Left            =   7200
         TabIndex        =   44
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txtBuatPIN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   43
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   41
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -74880
         TabIndex        =   32
         Top             =   360
         Width           =   8895
         Begin VB.Timer tmrCetak 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   2640
            Top             =   240
         End
         Begin VB.Timer tmrSimpan 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   3120
            Top             =   240
         End
         Begin VB.CommandButton Command1 
            Caption         =   "cek"
            Height          =   375
            Left            =   2040
            TabIndex        =   40
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdLihatPIN 
            Caption         =   "&Lihat Semua PIN"
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
            Left            =   4800
            TabIndex        =   37
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdUpload 
            Caption         =   "&Upload"
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
            Left            =   6240
            TabIndex        =   36
            Top             =   240
            Width           =   1215
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
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblFRS 
            AutoSize        =   -1  'True
            Caption         =   "<frs>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1320
            TabIndex        =   39
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label3 
            Caption         =   "Alamat FRS:"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid dgPIN 
         Height          =   2415
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4260
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   -74880
         TabIndex        =   31
         Top             =   1560
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PIN"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Alamat FRS"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FP"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama Lengkap"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "JK"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ruangan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Jabatan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tgl. Daftar"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cari Nama/PIN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -69480
         TabIndex        =   49
         Top             =   1200
         Width           =   1320
      End
      Begin VB.Label Label2 
         Caption         =   "Cari Pegawai/No. PIN:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   3360
         Width           =   2295
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   18
      Top             =   8715
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   16484
            Text            =   "F1 - Cetak"
            TextSave        =   "F1 - Cetak"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdtutup 
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
      Left            =   7080
      TabIndex        =   17
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Frame Frame2 
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   9135
      Begin VB.TextBox txtjk 
         Alignment       =   2  'Center
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtjabatan 
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
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   14
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtTempatTugas 
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
         Left            =   4680
         MaxLength       =   50
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtIDPegawai 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtNamaPegawai 
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
         Height          =   315
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   0
         Top             =   480
         Width           =   2295
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
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
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   3
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempat Bertugas"
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
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1230
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
         Left            =   6480
         TabIndex        =   4
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "PIN Absensi"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4095
      Begin VB.TextBox txttanggal 
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
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   16
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtNoPIN 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   120
         MaxLength       =   50
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Daftar"
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
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. PIN"
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
         TabIndex        =   11
         Top             =   240
         Width           =   555
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
   Begin VB.CommandButton cmdMinta 
      Caption         =   "&Minta PIN"
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
      TabIndex        =   19
      Top             =   8280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7560
      Picture         =   "frmPINAbsensiPegawai.frx":0D02
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   0
      Picture         =   "frmPINAbsensiPegawai.frx":1A8A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   1800
      Picture         =   "frmPINAbsensiPegawai.frx":444B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmPINAbsensiPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim msg As VbMsgBoxResult

Private Sub chkPilihSemua_Click()
    Dim itm As ListItem

    If Me.chkPilihSemua.Value = 1 Then
        For Each itm In Me.ListView1.ListItems
            itm.Checked = True
        Next
    Else
        For Each itm In Me.ListView1.ListItems
            itm.Checked = False
        Next
    End If
End Sub

Private Sub chkTidakPunyaPIN_Click()
    If Me.chkTidakPunyaPIN.Value = 1 Then
        LoadGridPIN True
    Else
        LoadGridPIN False
    End If
End Sub

Private Sub cmdBatal_Click()
    On Error GoTo errLoad
    Call clearData
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBuatPIN_Click()

    Me.tmrBuatPIN.Enabled = True
End Sub

Private Sub cmdCek_Click()

    If txtNoPIN.Text = "" Then
        MsgBox "Nomor PIN Tidak Ada !", vbExclamation, "Peringatan"
    Else
        If txtCek.Text = "" Then
            MsgBox "Alamat FRS400 Yang Akan Di Cek Kosong !", vbExclamation, "Peringatan"
        Else
            If frmAbsensiPegawai.MSComm1.PortOpen = True Then
                If txtCek.Text <= 0 Or txtCek.Text > add2 Then
                    MsgBox "Alamat FRS400 Asal atau Tujuan Salah !", vbExclamation, "Peringatan"
                Else
                    Dim cekSumDec As Integer
                    Dim cekSumHex As String
                    Dim protokolPIN As String, fRS As String
                    Dim nChr As Integer

                    frmAbsensiPegawai.minta_absensi.Enabled = False
                    fRS = Me.txtCek.Text
                    pinCek = ValidPIN(Me.txtNoPIN.Text)

                    cekSumDec = 11 Xor CInt(fRS) Xor cekSum(pinCek)
                    cekSumHex = Hex$(cekSumDec)

                    protokolPIN = Chr$(&H2) & Chr$(&HD) & Chr$(fRS) & Chr$(&H7) & _
                    pinCek & Chr$(3) & Chr$(cekSumDec)

                    frmAbsensiPegawai.MSComm1.Output = protokolPIN
                End If
            Else
                MsgBox "Tidak Ada Koneksi Ke FRS-400 !", vbExclamation, "Peringatan"
            End If
        End If
    End If

End Sub

Private Sub cmdHapusPIN_Click()

    dbConn.Execute "DELETE DataPegawai where IdPegawai = '" & txtIDPegawai.Text & "'"

    If txtNoPIN.Text = "" Then
        MsgBox "PIN Tidak Ada !", vbOKOnly, "Pesan Hapus PIN"
    Else

        i = MsgBox("Hapus PIN pada semua FRS-400!", vbOKCancel, "Hapus PIN")

        If i = vbOK Then
            dbConn.Execute "DELETE PINAbsensiPegawai where PINAbsensi = '" & txtNoPIN.Text & "' AND TglDaftar = '" & txttanggal.Text & "'"
            With frmAbsensiPegawai
                .timerPIN.Enabled = False
                .minta_absensi.Enabled = False
                frsHapus = &H0
                pinHapus = "00000000" & frmPINAbsensiPegawai.txtNoPIN.Text
                pinHapus = Right(pinHapus, 8)
                h11 = Asc(Mid(pinHapus, 1, 1)) 'diganti
                h12 = Asc(Mid(pinHapus, 2, 1)) 'diganti
                h13 = Asc(Mid(pinHapus, 3, 1)) 'diganti
                h14 = Asc(Mid(pinHapus, 4, 1)) 'diganti
                h15 = Asc(Mid(pinHapus, 5, 1)) 'diganti
                h16 = Asc(Mid(pinHapus, 6, 1)) 'diganti
                h17 = Asc(Mid(pinHapus, 7, 1)) 'diganti
                h18 = Asc(Mid(pinHapus, 8, 1)) 'diganti
                .TimerHapusPIN.Enabled = True
            End With

            Call clearData
            Call LoadGridPIN
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdHapusPinFRS_Click()
    Dim pesan As VbMsgBoxResult
    Dim cmdHapusPIN As ADODB.Command

    pinHapusFRS = ValidPIN(Me.txtNoPIN.Text)
    If Me.chkHapusDB.Value = 1 Then
        pesan = MsgBox("Hapus nomor PIN " & Me.txtNoPIN.Text & " dari data base?", vbInformation Or vbYesNo, "Hapus PIN")
        If pesan = vbYes Then
            Set cmdHapusPIN = New ADODB.Command
            With cmdHapusPIN
                .ActiveConnection = dbConn
                .CommandText = "DELETE PINAbsensiPegawai WHERE PINAbsensi=" & _
                funcPrepareString(Me.txtNoPIN.Text)
                .CommandType = adCmdText
                .Execute
            End With
            LoadGridPIN
        End If
    Else
        frsHapus = CInt(Me.txtCek.Text)
        pesan = MsgBox("Hapus PIN: " & pinHapusFRS & " pada FRS-" & frsHapus & "?", vbQuestion Or vbYesNo, "Hapus PIN..")
        If pesan = vbYes Then
            frmAbsensiPegawai.minta_absensi.Enabled = False
            protokolHapusPin = Chr$(&H2) & Chr$(&HD) & Chr$(frsHapus) & Chr$(&H8) & _
            pinHapusFRS & Chr$(&H3)
            protokolHapusPin = protokolHapusPin & Chr$(cekSum(protokolHapusPin))
            frmAbsensiPegawai.MSComm1.Output = protokolHapusPin
            statusPinHapus = True
            LoadGridPIN
        End If
    End If
End Sub

Private Sub cmdLihatPIN_Click()
    If frmAbsensiPegawai.MSComm1.PortOpen Then
        strStatusSekarang = "prepare"
        subCekJumlahPIN Me.txtCek.Text
    Else
        MsgBox Error$, 48, "Peringatan"
    End If
End Sub

Private Sub cmdReset_Click()
    Dim pesan As VbMsgBoxResult

    If Me.SSTab1.Tab = 0 Then
        pesan = MsgBox("Reset alamat PIN?", vbInformation Or vbYesNo, "Perhatian")
        If pesan = vbYes Then
            strSQL = "UPDATE PINAbsensiPegawai SET AlamatFRS=NULL"
            Set rs = New ADODB.recordset
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
        End If
        LoadGridPIN
    ElseIf Me.SSTab1.Tab = 1 Then
        Me.ListView1.ListItems.clear
    End If
End Sub

Private Sub cmdSimpan_Click()
    If Me.ListView1.ListItems.Count = 0 Then
        MsgBox "Daftar PIN yang akan disimpan kosong!", vbExclamation, "Perhatian"
        Exit Sub
    End If

    Dim itm As ListItem
    Dim itmChecked As Boolean

    For Each itm In Me.ListView1.ListItems
        If itm.Checked Then
            itmChecked = True
            Exit For
        Else
            itmChecked = False
        End If
    Next
    If Not itmChecked Then
        MsgBox "Tidak ada PIN yang dipilih!", vbExclamation, "Perhatian"
        Exit Sub
    End If

    frmAbsensiPegawai.minta_absensi.Enabled = False
    frmAbsensiPegawai.tmr_CekError.Enabled = False
    With frmStatusProses
        .lblStatus.Caption = "Simpan PIN"
        .lblPIN.Caption = ""
        .pgbStatus.Max = Me.ListView1.ListItems.Count
        .pgbStatus.Min = 0
        .pgbStatus.Value = 0
        .Show
    End With
    strStatusSekarang = "simpan"
    Me.tmrSimpan.Enabled = True
End Sub

Private Sub cmdSimpanPIN_Click()
    Dim pesan As VbMsgBoxResult
    Dim sqlGantiPIN As String
    Dim cmdGantiPIN As ADODB.Command

    If Me.txtIDPegawai.Text = "" Then
        MsgBox "Tidak ada ID Pegawai yang dipilih!" & vbNewLine & _
        "Silahkan pilih ID Pegawai dari daftar yang ada.", vbExclamation, "Perhatian"
        Exit Sub
    End If
    If Me.txtNoPIN.Text <> "" Then
        pesan = MsgBox("Ganti PIN untuk " & Me.txtNamaPegawai.Text & " dari nomor PIN " & Me.txtNoPIN.Text & _
        " ke nomor PIN " & Me.txtBuatPIN.Text & "?", vbInformation Or vbYesNo, "Ganti PIN")
        If pesan = vbYes Then
            sqlGantiPIN = "UPDATE PINAbsensiPegawai SET" & _
            " PINAbsensi=" & funcPrepareString(Me.txtBuatPIN.Text) & _
            " WHERE IdPegawai= " & funcPrepareString(Me.txtIDPegawai.Text)

            Set cmdGantiPIN = New ADODB.Command

            With cmdGantiPIN
                .ActiveConnection = dbConn
                .CommandText = sqlGantiPIN
                .CommandType = adCmdText
                .Execute
            End With
            LoadGridPIN
        End If
    Else
        pesan = MsgBox("Tambah PIN untuk " & Me.txtNamaPegawai.Text & " dengan nomor PIN: " & _
        Me.txtBuatPIN.Text & "?", vbInformation Or vbYesNo, "Tambah PIN")
        If pesan = vbYes Then
            sqlGantiPIN = "INSERT INTO PINAbsensiPegawai (IdPegawai,PINAbsensi) VALUES (" & _
            funcPrepareString(Me.txtIDPegawai.Text) & "," & _
            funcPrepareString(Me.txtBuatPIN.Text) & ")"

            Set cmdGantiPIN = New ADODB.Command

            With cmdGantiPIN
                .ActiveConnection = dbConn
                .CommandText = sqlGantiPIN
                .CommandType = adCmdText
                .Execute
            End With
            LoadGridPIN
        End If
    End If
End Sub

Private Sub cmdTransfer_Click()
    Dim protokolUpload As String

    If txtNoPIN.Text = "" Then
        MsgBox "Nomor PIN Tidak Ada !", vbExclamation, "Peringatan"
    Else
        If add2 = "" Then
            MsgBox "Belum ada koneksi dengan FRS400.", vbExclamation, "Perhatian"
            Exit Sub
        End If
        If Me.txtTujuan.Text = "" Then
            MsgBox "Alamat FRS tujuan harus diisi!", vbExclamation, "Peringatan"
            Exit Sub
        ElseIf CInt(Me.txtTujuan.Text) <= 0 Or CInt(Me.txtTujuan.Text) > add2 Then
            MsgBox "Alamat FRS400 asal atau tujuan salah!", vbExclamation, "Peringatan"
            Exit Sub
        End If
        frmAbsensiPegawai.minta_absensi.Enabled = False
        frmAbsensiPegawai.tmr_CekError.Enabled = False
        frsTujuan = CInt(Me.txtTujuan.Text)
        pinCek = ValidPIN(Me.txtNoPIN.Text)
        Dim protokolDownload As String

        dl = 1
        Set rs = New ADODB.recordset
        rs.Open "SELECT FingerPrint1 FROM PINAbsensiPegawai " & _
        "WHERE PINAbsensi=" & funcPrepareString(CInt(pinCek)), dbConn, 3, 3
        Dim strImageData As String
        If Not IsNull(rs.Fields.Item(0).Value) Then
            strImageData = rs.Fields(0).Value
        Else
            MsgBox "Nomor PIN " & CInt(pinCek) & " belum memiliki data finger print!" & vbNewLine & _
            "Silahkan lakukan upload data terlebih dahulu.", vbExclamation, "Perhatian"
            frmAbsensiPegawai.minta_absensi.Enabled = True
            frmAbsensiPegawai.tmr_CekError.Enabled = True
            Exit Sub
        End If
        varImageData = funcConvertHexKeImage(strImageData)
        protokolDownload = Chr$(&H2) & Chr$(&HD7) & Chr$(frsTujuan) & Chr$(&H10) & _
        pinCek & Chr$(&H30) & Chr$(&H31) & varImageData & Chr$(&H3)
        protokolDownload = protokolDownload & Chr$(cekSum(protokolDownload))
        rs.Close
        statusTransfer = True
        statDownload = True
        frmAbsensiPegawai.MSComm1.Output = protokolDownload
    End If

End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub cmdUpload_Click()
    If Me.ListView1.ListItems.Count = 0 Then
        MsgBox "Daftar PIN yang akan di upload kosong!", vbCritical, "Perhatian.."
        Exit Sub
    End If

    Dim itm As ListItem
    Dim itmChecked As Boolean

    For Each itm In Me.ListView1.ListItems
        If itm.Checked Then
            itmChecked = True
            Exit For
        Else
            itmChecked = False
        End If
    Next
    If Not itmChecked Then
        MsgBox "Tidak ada PIN yang dipilih!", vbExclamation, "Perhatian"
        Exit Sub
    End If

    Dim protokolUpload As String

    frmAbsensiPegawai.minta_absensi.Enabled = False
    frmAbsensiPegawai.tmr_CekError.Enabled = False

    With frmStatusProses
        .lblStatus.Caption = "Upload PIN..."
        .lblPIN.Caption = ""
        .pgbStatus.Max = Me.ListView1.ListItems.Count
        .pgbStatus.Min = 0
        .pgbStatus.Value = 1
        .Show
    End With
balikmaning:
    idxListViewPIN = idxListViewPIN + 1
    If idxListViewPIN > Me.ListView1.ListItems.Count Then
        MsgBox "Upload semua PIN dari FRS-" & frsLihatPIN & " selesai!", vbInformation, "Perhatian"
        idxListViewPIN = 0
        bolFullUpload = False
        frmAbsensiPegawai.minta_absensi.Enabled = True
        frmAbsensiPegawai.tmr_CekError.Enabled = True
        Unload frmStatusProses
        Exit Sub
    End If
    frmStatusProses.pgbStatus.Value = idxListViewPIN
    If Not Me.ListView1.ListItems.Item(idxListViewPIN).Checked Then GoTo balikmaning
    If frmPINAbsensiPegawai.ListView1.ListItems.Item(idxListViewPIN).SubItems(5) = "" Then GoTo balikmaning
    pinSimpan = Me.ListView1.ListItems.Item(idxListViewPIN).Text
    frsLihatPIN = Me.txtCek.Text
    Dim crFRS As Integer
    crFRS = InStr(1, Me.ListView1.ListItems.Item(idxListViewPIN).SubItems(1), frsLihatPIN, vbTextCompare)
    If crFRS = 0 Then GoTo balikmaning
    Me.ListView1.ListItems.Item(idxListViewPIN).SubItems(2) = "Y"
    protokolUpload = funcBuatProtokolUpload(frsLihatPIN, pinSimpan)
    frmAbsensiPegawai.MSComm1.Output = protokolUpload
    bolFullUpload = True
    strStatusSekarang = "fullupload"
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Open App.Path & "\daftarpin.txt" For Output As #1
        For i = 1 To jumlahTotalPIN
            Me.ListView1.ListItems.add , , tempDataPIN(i)
            Print #1, tempDataPIN(i)
        Next
    Close #1
    MsgBox Me.ListView1.ListItems.Count
End Sub

Private Sub dgPIN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgPIN
        txtIDPegawai.Text = .Columns("ID").Value
        txtNamaPegawai.Text = .Columns("Nama").Value
        txtjk.Text = .Columns("JK").Value
        txtTempatTugas.Text = .Columns("Ruangan").Value
        txtjabatan.Text = .Columns("Jabatan").Value
        If .Columns("ID FP").Value = "" Then
            txtNoPIN.Text = ""
        Else
            txtNoPIN.Text = .Columns("ID FP").Value
        End If
        If .Columns("Tgl. Daftar").Value = "" Then
            txttanggal.Text = ""
        Else
            txttanggal.Text = .Columns("Tgl. Daftar").Value
        End If
        mstrIdPegawai = txtIDPegawai.Text
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    centerForm Me, MDIUtama
    Call cmdBatal_Click
    Call PlayFlashMovie(Me)
    Call LoadGridPIN
    Me.lblFRS.Caption = ""
    intBuatPIN = 1
    Me.SSTab1.Tab = 0
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub clearData()
    On Error Resume Next
    txtIDPegawai.Text = ""
    txtNamaPegawai.Text = ""
    txtjk.Text = ""
    txtTempatTugas.Text = ""
    txtjabatan.Text = ""
    txtNoPIN.Text = ""
    txttanggal.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        frmCetakPINAbsensi.Show
    End If
End Sub

Public Function sp_PIN() As Boolean
    sp_PIN = True
    Set adoComm = New ADODB.Command
    With adoComm
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("PINAbsensi", adVarChar, adParamInput, 8, pinSimpan)
        .Parameters.Append .CreateParameter("AlamatFRS", adVarChar, adParamInput, 5, alamatFRS)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, mstrIdPegawai)
        .Parameters.Append .CreateParameter("TglDaftar", adDate, adParamInput, , IIf(txttanggal.Text = "", Null, Format(txttanggal.Text, "yyyy/MM/dd")))

        .ActiveConnection = dbConn
        .CommandText = "AU_PINAbsensi"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
        End If
        Call deleteADOCommandParameters(adoComm)
        Set adoComm = Nothing
    End With
End Function

Private Sub LoadGridPIN(Optional ByVal unPIN As Boolean)
    On Error GoTo errLoad
    If Not unPIN Then
        strSQL = "SELECT * FROM v_PIN"
    Else
        strSQL = "SELECT * FROM v_PIN WHERE PIN IS NULL"
    End If
    dbConn.Execute strSQL
    Call msubRecFO(rs, strSQL)
    Set dgPIN.DataSource = rs
    dgPIN.Columns("Tgl. Mulai").Width = 0
    Exit Sub
errLoad:
End Sub

Private Sub cmdMinta_Click()

    If txtNoPIN.Text = "" And txtIDPegawai.Text <> "" Then
        frmKPIN.Show , Me

    ElseIf txtNoPIN.Text <> "" Then
        MsgBox "PIN Sudah Ada, Tidak Bisa Minta PIN !", vbOKOnly, "Pesan Minta PIN"
    ElseIf txtIDPegawai.Text = "" Then
        MsgBox "ID Pegawai Tidak Ada, Tidak Bisa Minta PIN !", vbOKOnly, "Pesan Minta PIN"
    End If
End Sub

Private Sub ListView1_DblClick()
    If Me.ListView1.ListItems.Count = 0 Then Exit Sub
    If Me.ListView1.SelectedItem.SubItems(4) <> "" Then
        MsgBox "PIN sudah memiliki ID.", vbExclamation, "Perhatian"
        Exit Sub
    End If
    idxListViewPIN = Me.ListView1.SelectedItem.Index
    frmPIlihID.Show
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.txtIDPegawai.Text = Item.SubItems(5)
    Me.txtNamaPegawai.Text = Item.SubItems(3)
    Me.txtjabatan.Text = Item.SubItems(7)
    Me.txtjk.Text = Item.SubItems(4)
    Me.txtNoPIN.Text = Item.Text
    Me.txttanggal.Text = Item.SubItems(8)
    Me.txtTempatTugas.Text = Item.SubItems(6)
End Sub

Private Sub tmrBuatPIN_Timer()
    Dim sql As String
    sql = "SELECT ID FROM v_PIN WHERE PIN='" & intBuatPIN & "'"
    Me.txtBuatPIN.Text = intBuatPIN
    Set rs = New ADODB.recordset
    rs.Open sql, dbConn, adOpenForwardOnly, adLockReadOnly
    If rs.RecordCount = 0 Then
        intBuatPIN = intBuatPIN + 1
        Me.tmrBuatPIN.Enabled = False
    End If
    rs.Close
    intBuatPIN = intBuatPIN + 1
End Sub

Private Sub tmrCetak_Timer()
    Static i As Integer
    Dim N As Integer
    Static unID As Integer
    Dim j As Integer
    Dim cr As Integer, sama As Integer
    Dim lblStat As Label, lblPIN As Label
    Dim pgb As ProgressBar
    Dim lsv As ListView
    Dim rsFP As ADODB.recordset
    Dim sqlFP As String

    If resetInteger Then
        i = 0
        unID = 0
        resetInteger = False
    End If

    With frmStatusProses
        Set lblStat = .lblStatus
        Set lblPIN = .lblPIN
        Set pgb = .pgbStatus
    End With
    Set lsv = Me.ListView1

    lblStat.Caption = "Cetak PIN"
    pgb.Max = jumlahTotalPIN
    pgb.Min = 0
    i = i + 1
    If i > jumlahTotalPIN Then
        i = 0
        Me.tmrCetak.Enabled = False
        GoTo Selesai
    End If
    lblPIN.Caption = tempDataPIN(i)
    strSQL = "SELECT * FROM v_PIN WHERE PIN=" & "'" & tempDataPIN(i) & "'"
    dbConn.Execute strSQL
    Call msubRecFO(rs, strSQL)
    If lsv.ListItems.Count = 0 Then
        If rs.RecordCount = 0 Then
            lsv.ListItems.add(, , tempDataPIN(i)).SubItems(1) = frsTujuan
            unID = unID + 1
        Else
            With lsv.ListItems.add(, , tempDataPIN(i))
                sqlFP = "SELECT FingerPrint4 FROM PINAbsensiPegawai WHERE PINAbsensi='" & tempDataPIN(i) & "'"
                Set rsFP = New ADODB.recordset
                rsFP.Open sqlFP, dbConn, adOpenForwardOnly, adLockReadOnly
                If IsNull(rsFP.Fields.Item(0).Value) Then
                    .SubItems(2) = "T"
                Else
                    .SubItems(2) = "Y"
                End If
                rsFP.Close
                .SubItems(1) = frsTujuan
                .SubItems(3) = IIf(IsNull(rs.Fields.Item("Nama").Value), "", rs.Fields.Item("Nama").Value)
                .SubItems(4) = IIf(IsNull(rs.Fields.Item("JK").Value), "", rs.Fields.Item("JK").Value)
                .SubItems(5) = IIf(IsNull(rs.Fields.Item("ID").Value), "", rs.Fields.Item("ID").Value)
                .SubItems(6) = IIf(IsNull(rs.Fields.Item("Ruangan").Value), "", rs.Fields.Item("Ruangan").Value)
                .SubItems(7) = IIf(IsNull(rs.Fields.Item("Jabatan").Value), "", rs.Fields.Item("Jabatan").Value)
                .SubItems(8) = IIf(IsNull(rs.Fields.Item("Tgl. Daftar").Value), "", rs.Fields.Item("Tgl. Daftar").Value)
            End With
        End If
    Else
        For j = 1 To lsv.ListItems.Count
            If tempDataPIN(i) = lsv.ListItems.Item(j).Text Then
                cr = InStr(1, lsv.ListItems(j).SubItems(1), frsTujuan, vbTextCompare)
                If cr = 0 Then
                    lsv.ListItems.Item(j).SubItems(1) = lsv.ListItems.Item(j).SubItems(1) & "," & frsTujuan
                End If
                sama = sama + 1
                GoTo lompat
            End If
        Next
        If rs.RecordCount = 0 Then
            lsv.ListItems.add(, , tempDataPIN(i)).SubItems(1) = frsTujuan
            unID = unID + 1
        Else
            With lsv.ListItems.add(, , tempDataPIN(i))
                sqlFP = "SELECT FingerPrint4 FROM PINAbsensiPegawai WHERE PINAbsensi='" & tempDataPIN(i) & "'"
                Set rsFP = New ADODB.recordset
                rsFP.Open sqlFP, dbConn, adOpenForwardOnly, adLockReadOnly
                If IsNull(rsFP.Fields.Item(0).Value) Then
                    .SubItems(2) = "T"
                Else
                    .SubItems(2) = "Y"
                End If
                rsFP.Close
                .SubItems(1) = frsTujuan
                .SubItems(3) = IIf(IsNull(rs.Fields.Item("Nama").Value), "", rs.Fields.Item("Nama").Value)
                .SubItems(4) = IIf(IsNull(rs.Fields.Item("JK").Value), "", rs.Fields.Item("JK").Value)
                .SubItems(5) = IIf(IsNull(rs.Fields.Item("ID").Value), "", rs.Fields.Item("ID").Value)
                .SubItems(6) = IIf(IsNull(rs.Fields.Item("Ruangan").Value), "", rs.Fields.Item("Ruangan").Value)
                .SubItems(7) = IIf(IsNull(rs.Fields.Item("Jabatan").Value), "", rs.Fields.Item("Jabatan").Value)
                .SubItems(8) = IIf(IsNull(rs.Fields.Item("Tgl. Daftar").Value), "", rs.Fields.Item("Tgl. Daftar").Value)
            End With
        End If
lompat:
    End If
    N = N + 1
    pgb.Value = i
    Exit Sub
Selesai:
    frsLihatPIN = frsTujuan
    If unID > 0 Then
        MsgBox "Jumlah PIN yang tidak memiliki data kepemilikan: " & unID & " dari " & jumlahTotalPIN & " pin.", vbExclamation, "Perhatian"
    End If
    frmAbsensiPegawai.minta_absensi.Enabled = True
    frmAbsensiPegawai.tmr_CekError.Enabled = True
    Unload frmStatusProses
End Sub

Private Sub tmrSimpan_Timer()
    On Error GoTo ErrorHandling
    Dim N As Integer
    Static c As Integer

    N = Me.ListView1.ListItems.Count
    If resetInteger Then
        c = 0
        resetInteger = False
    End If
    c = c + 1
    With frmStatusProses
        alamatFRS = Me.ListView1.ListItems.Item(c).SubItems(1)
        pinSimpan = Me.ListView1.ListItems.Item(c).Text
        .lblStatus.Caption = "Simpan PIN " & pinSimpan
        If Me.ListView1.ListItems.Item(c).SubItems(3) <> "" Then
            If Not Me.ListView1.ListItems.Item(c).Checked Then GoTo jump
            strSQL = "SELECT PINAbsensi FROM PINAbsensiPegawai WHERE IdPegawai=" & funcPrepareString(Me.ListView1.ListItems.Item(c).SubItems(5))
            Set rs = New ADODB.recordset
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount = 0 Then
                Set adoComm = New ADODB.Command
                adoComm.ActiveConnection = dbConn
                adoComm.CommandText = "INSERT INTO PINAbsensiPegawai " & _
                "(IdPegawai,PINAbsensi,AlamatFRS) VALUES (" & _
                funcPrepareString(Me.ListView1.ListItems.Item(c).SubItems(5)) & "," & _
                funcPrepareString(Me.ListView1.ListItems.Item(c).Text) & "," & _
                funcPrepareString(alamatFRS) & ")"
                adoComm.CommandType = adCmdText
                adoComm.Execute
            Else
                Set adoComm = New ADODB.Command
                adoComm.ActiveConnection = dbConn
                adoComm.CommandText = "UPDATE PINAbsensiPegawai SET " & _
                "AlamatFRS=" & funcPrepareString(alamatFRS) & " " & _
                "WHERE IdPegawai=" & funcPrepareString(Me.ListView1.ListItems.Item(c).SubItems(5))
                adoComm.CommandType = adCmdText
                adoComm.Execute
            End If
            rs.Close
        End If
jump:
        .pgbStatus.Value = c
    End With
    If c = N Then
        Me.tmrSimpan.Enabled = False
        c = 0
        MsgBox "Simpan PIN selesai.", vbInformation, "Perhatian"
        LoadGridPIN
        frmAbsensiPegawai.minta_absensi.Enabled = True
        frmAbsensiPegawai.tmr_CekError.Enabled = True
        Unload frmStatusProses
    End If
    Exit Sub
ErrorHandling:
    msubPesanError
End Sub

Private Sub txtBuatPIN_Change()
    If Me.txtBuatPIN.Text = "" Then
        Me.cmdSimpanPIN.Enabled = False
    Else
        Me.cmdSimpanPIN.Enabled = True
    End If
End Sub

Private Sub txtBuatPIN_DblClick()
    If Me.tmrBuatPIN.Enabled Then Exit Sub
    intBuatPIN = 1
    Me.txtBuatPIN.Text = ""
End Sub

Private Sub txtCari_Change()
    Dim itm As ListItem
    Dim cr1 As Integer, cr2 As Integer

    For Each itm In Me.ListView1.ListItems
        cr1 = InStr(1, itm.Text, Me.txtCari.Text, vbTextCompare)
        cr2 = InStr(1, itm.SubItems(3), Me.txtCari.Text, vbTextCompare)
        If cr1 > 0 Or cr2 > 0 Then
            itm.Selected = True
            itm.EnsureVisible
            Exit For
        End If
    Next
End Sub

Private Sub txtCek_Change()
    Me.lblFRS.Caption = Me.txtCek.Text
End Sub

Private Sub txtNoPIN_Change()
    If Len(Trim(Me.txtNoPIN.Text)) = 0 Then
        Me.cmdHapusPinFRS.Enabled = False
    Else
        Me.cmdHapusPinFRS.Enabled = True
    End If
End Sub

Private Sub txtParameter_Change()
    Set rs = Nothing
    strSQL = "SELECT * FROM v_PIN WHERE ID LIKE '%" & txtParameter.Text & "%' OR Nama LIKE '%" & txtParameter.Text & "%' OR PIN LIKE '%" & txtParameter.Text & "%'"
    dbConn.Execute strSQL
    Call msubRecFO(rs, strSQL)
    Set dgPIN.DataSource = rs
End Sub
