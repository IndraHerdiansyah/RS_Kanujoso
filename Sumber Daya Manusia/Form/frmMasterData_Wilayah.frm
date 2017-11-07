VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMasterWilayah 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Wilayah"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterData_Wilayah.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7980
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   5160
      TabIndex        =   21
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdTutup 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   6480
      TabIndex        =   22
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab sstWilayah 
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Propinsi"
      TabPicture(0)   =   "frmMasterData_Wilayah.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Kota/Kabupaten"
      TabPicture(1)   =   "frmMasterData_Wilayah.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame8"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Kecamatan"
      TabPicture(2)   =   "frmMasterData_Wilayah.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Kelurahan/Desa"
      TabPicture(3)   =   "frmMasterData_Wilayah.frx":0D1E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   7455
         Begin VB.CheckBox chkSts 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4440
            TabIndex        =   61
            Top             =   2160
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtNmExt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   60
            Top             =   2520
            Width           =   4215
         End
         Begin VB.TextBox txtKdExt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   59
            Top             =   2160
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo dc4Kecamatan 
            Height          =   315
            Left            =   240
            TabIndex        =   14
            Top             =   1125
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataListLib.DataCombo dc4KotaKabupaten 
            Height          =   315
            Left            =   3930
            TabIndex        =   13
            Top             =   480
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin VB.TextBox txt4KodePos 
            Alignment       =   2  'Center
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
            Height          =   330
            Left            =   5400
            MaxLength       =   10
            TabIndex        =   16
            Top             =   1125
            Width           =   1815
         End
         Begin VB.TextBox txt4KelurahanDesa 
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
            Height          =   330
            Left            =   240
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1710
            Width           =   7005
         End
         Begin VB.TextBox txt4KdKelurahanDesa 
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
            Height          =   330
            Left            =   3840
            MaxLength       =   9
            TabIndex        =   15
            Top             =   1125
            Width           =   1455
         End
         Begin MSDataListLib.DataCombo dc4Propinsi 
            Height          =   315
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataGridLib.DataGrid dgKelurahanDesa 
            Height          =   2475
            Left            =   240
            TabIndex        =   18
            Top             =   2880
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   4366
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   195
            Index           =   7
            Left            =   255
            TabIndex        =   63
            Top             =   2535
            Width           =   1275
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   62
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Kode Pos"
            Height          =   210
            Left            =   5400
            TabIndex        =   42
            Top             =   885
            Width           =   765
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nama Kecamatan"
            Height          =   210
            Left            =   255
            TabIndex        =   41
            Top             =   885
            Width           =   1410
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kel. / Desa"
            Height          =   210
            Left            =   3840
            TabIndex        =   40
            Top             =   885
            Width           =   1365
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nama Kota / Kabupaten"
            Height          =   210
            Left            =   3960
            TabIndex        =   39
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nama Propinsi"
            Height          =   210
            Left            =   240
            TabIndex        =   38
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nama Kelurahan / Desa"
            Height          =   210
            Left            =   240
            TabIndex        =   37
            Top             =   1470
            Width           =   1890
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   31
         Top             =   480
         Width           =   7455
         Begin VB.CheckBox chkStsKecamatan 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4320
            TabIndex        =   56
            Top             =   1560
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtNmExtKecamatan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   55
            Top             =   1920
            Width           =   4215
         End
         Begin VB.TextBox txtKdExtKecamatan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   54
            Top             =   1560
            Width           =   2295
         End
         Begin MSDataListLib.DataCombo dc3KotaKabupaten 
            Height          =   315
            Left            =   3960
            TabIndex        =   23
            Top             =   495
            Width           =   3240
            _ExtentX        =   5715
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin VB.TextBox txt3KdKecamatan 
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
            Height          =   330
            Left            =   240
            MaxLength       =   6
            TabIndex        =   0
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txt3Kecamatan 
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
            Height          =   330
            Left            =   1785
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1080
            Width           =   5415
         End
         Begin MSDataListLib.DataCombo dc3Propinsi 
            Height          =   315
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataGridLib.DataGrid dgKecamatan 
            Height          =   2610
            Left            =   240
            TabIndex        =   11
            Top             =   2400
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   4604
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   195
            Index           =   5
            Left            =   255
            TabIndex        =   58
            Top             =   1935
            Width           =   1515
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   57
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nama Kecamatan"
            Height          =   210
            Left            =   1800
            TabIndex        =   35
            Top             =   840
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nama Propinsi"
            Height          =   210
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nama Kota / Kabupaten"
            Height          =   210
            Left            =   3960
            TabIndex        =   33
            Top             =   240
            Width           =   1965
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kecamatan"
            Height          =   210
            Left            =   240
            TabIndex        =   32
            Top             =   840
            Width           =   1380
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5175
         Left            =   -74760
         TabIndex        =   25
         Top             =   480
         Width           =   7455
         Begin VB.CheckBox chkStsPropinsi 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4440
            TabIndex        =   46
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtNmExtPropinsi 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox txtKdExtPropinsi 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   44
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txtNamaPropinsi 
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
            Height          =   330
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   3
            Top             =   600
            Width           =   5175
         End
         Begin VB.TextBox txtKdPropinsi 
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
            Height          =   330
            Left            =   240
            MaxLength       =   2
            TabIndex        =   2
            Top             =   600
            Width           =   1095
         End
         Begin MSDataGridLib.DataGrid dgPropinsi 
            Height          =   3105
            Left            =   240
            TabIndex        =   4
            Top             =   1920
            Width           =   6960
            _ExtentX        =   12277
            _ExtentY        =   5477
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   195
            Index           =   3
            Left            =   255
            TabIndex        =   48
            Top             =   1455
            Width           =   1275
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nama Propinsi"
            Height          =   210
            Left            =   1680
            TabIndex        =   27
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Kode Propinsi"
            Height          =   210
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Height          =   5175
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   7455
         Begin VB.CheckBox chkStsKota 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4320
            TabIndex        =   51
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox txtNmExtKota 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   50
            Top             =   1440
            Width           =   4215
         End
         Begin VB.TextBox txtKdExtKota 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   49
            Top             =   1080
            Width           =   2295
         End
         Begin VB.TextBox txt2KotaKabupaten 
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
            Height          =   330
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   7
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txt2KdKotaKabupaten 
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
            Height          =   330
            Left            =   3360
            MaxLength       =   4
            TabIndex        =   6
            Top             =   600
            Width           =   855
         End
         Begin MSDataListLib.DataCombo dc2Propinsi 
            Height          =   315
            Left            =   240
            TabIndex        =   5
            Top             =   600
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
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
         Begin MSDataGridLib.DataGrid dgKotaKabupaten 
            Height          =   3045
            Left            =   240
            TabIndex        =   8
            Top             =   1920
            Width           =   7020
            _ExtentX        =   12383
            _ExtentY        =   5371
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   195
            Index           =   1
            Left            =   255
            TabIndex        =   53
            Top             =   1455
            Width           =   1155
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   52
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nama Kota/Kabupaten"
            Height          =   210
            Left            =   4320
            TabIndex        =   30
            Top             =   360
            Width           =   1845
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Kode Kab."
            Height          =   210
            Left            =   3360
            TabIndex        =   29
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nama Propinsi"
            Height          =   210
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Width           =   1125
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   43
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
      Left            =   6120
      Picture         =   "frmMasterData_Wilayah.frx":0D3A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterData_Wilayah.frx":1AC2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterData_Wilayah.frx":3120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmMasterWilayah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCommand As New ADODB.Command

Private Sub subLoadGridSource()
On Error GoTo errLoad
    Select Case sstWilayah.Tab
        Case 0 'Propinsi
            strsql = "SELECT * FROM Propinsi"
            Call msubRecFO(rs, strsql)
            Set dgPropinsi.DataSource = rs
                dgPropinsi.Columns(0).Width = 1500
                dgPropinsi.Columns(1).Width = 5000
                dgPropinsi.ReBind
        
        Case 1 'KotaKabupaten
            strsql = "SELECT dbo.Propinsi.NamaPropinsi, dbo.KotaKabupaten.KdPropinsi, dbo.KotaKabupaten.KdKotaKabupaten, dbo.KotaKabupaten.NamaKotaKabupaten, " & _
                " dbo.KotaKabupaten.KodeExternal, dbo.KotaKabupaten.NamaExternal, dbo.KotaKabupaten.StatusEnabled FROM dbo.KotaKabupaten INNER JOIN dbo.Propinsi ON dbo.KotaKabupaten.KdPropinsi = dbo.Propinsi.KdPropinsi" & _
                " ORDER BY dbo.Propinsi.NamaPropinsi, dbo.KotaKabupaten.NamaKotaKabupaten"
            Call msubRecFO(rs, strsql)
            Set dgKotaKabupaten.DataSource = rs
                dgKotaKabupaten.Columns(0).DataField = rs(0).Name 'propinsi
                dgKotaKabupaten.Columns(0).Width = 2500
                dgKotaKabupaten.Columns(1).DataField = rs(1).Name 'kode propinsi
                dgKotaKabupaten.Columns(2).DataField = rs(2).Name 'kode kota/kabupaten
                dgKotaKabupaten.Columns(3).DataField = rs(3).Name 'kota/kabupaten
                dgKotaKabupaten.Columns(3).Width = 4000
                dgKotaKabupaten.ReBind
        
        Case 2  'Kecamatan
            strsql = "SELECT dbo.Propinsi.NamaPropinsi, dbo.Kecamatan.KdPropinsi, dbo.Kecamatan.KdKotaKabupaten, dbo.KotaKabupaten.NamaKotaKabupaten, dbo.Kecamatan.KdKecamatan, dbo.Kecamatan.NamaKecamatan, " & _
                " dbo.Kecamatan.KodeExternal, dbo.Kecamatan.NamaExternal, dbo.Kecamatan.StatusEnabled FROM  dbo.KotaKabupaten INNER JOIN dbo.Kecamatan ON dbo.KotaKabupaten.KdPropinsi = dbo.Kecamatan.KdPropinsi AND dbo.KotaKabupaten.KdKotaKabupaten = dbo.Kecamatan.KdKotaKabupaten INNER JOIN dbo.Propinsi ON dbo.KotaKabupaten.KdPropinsi = dbo.Propinsi.KdPropinsi" & _
                " ORDER BY dbo.Propinsi.NamaPropinsi, dbo.KotaKabupaten.NamaKotaKabupaten, dbo.Kecamatan.NamaKecamatan"
            Call msubRecFO(rs, strsql)
            Set dgKecamatan.DataSource = rs
                dgKecamatan.Columns(0).Width = 2500
                dgKecamatan.Columns(3).Width = 2500
                dgKecamatan.Columns(5).Width = 3000
                dgKecamatan.ReBind
           
        Case 3  'KelurahanDesa
            strsql = "SELECT * FROM V_MasterWilayah"
'            strSQL = "SELECT dbo.Propinsi.NamaPropinsi, dbo.Kelurahan.KdPropinsi, dbo.Kelurahan.KdKotaKabupaten, dbo.KotaKabupaten.NamaKotaKabupaten, dbo.Kelurahan.KdKecamatan, dbo.Kecamatan.NamaKecamatan, dbo.Kelurahan.KdKelurahan, dbo.Kelurahan.KodePos, dbo.Kelurahan.NamaKelurahan" & _
'                " FROM dbo.KotaKabupaten INNER JOIN dbo.Kecamatan ON dbo.KotaKabupaten.KdPropinsi = dbo.Kecamatan.KdPropinsi AND dbo.KotaKabupaten.KdKotaKabupaten = dbo.Kecamatan.KdKotaKabupaten INNER JOIN dbo.Propinsi ON dbo.KotaKabupaten.KdPropinsi = dbo.Propinsi.KdPropinsi INNER JOIN dbo.Kelurahan ON dbo.Kecamatan.KdPropinsi = dbo.Kelurahan.KdPropinsi AND dbo.Kecamatan.KdKotaKabupaten = dbo.Kelurahan.KdKotaKabupaten AND dbo.Kecamatan.KdKecamatan = dbo.Kelurahan.KdKecamatan" & _
'                " ORDER BY dbo.Propinsi.NamaPropinsi, dbo.KotaKabupaten.NamaKotaKabupaten, dbo.Kecamatan.NamaKecamatan, dbo.Kelurahan.NamaKelurahan"
            Call msubRecFO(rs, strsql)
            Set dgKelurahanDesa.DataSource = rs
                dgKelurahanDesa.ReBind
    End Select
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdBatal_Click()
On Error GoTo errLoad
    Call clear
    Call subLoadDcSource
    Call subLoadGridSource
    Call sstWilayah_KeyPress(13)
    Call openConnection
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdHapus_Click()
On Error GoTo errLoad
    
    If MsgBox("Apakah anda yakin akan mengapus data ini", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub

    Select Case sstWilayah.Tab
        Case 0 'Propinsi
            If txtKdPropinsi.Text = "" Then Exit Sub
            If sp_Propinsi("D") = False Then Exit Sub
        
        Case 1 'KotaKabupaten
            If txt2KdKotaKabupaten.Text = "" Then Exit Sub
            If sp_KotaKabupaten("D") = False Then Exit Sub
        
        Case 2 'Kecamatan
            If txt3KdKecamatan.Text = "" Then Exit Sub
            If sp_Kecamatan("D") = False Then Exit Sub
        
        Case 3 'KelurahanDesa
            If txt4KdKelurahanDesa.Text = "" Then Exit Sub
            If sp_Kelurahan("D") = False Then Exit Sub
    
    End Select

    MsgBox "Data berhasil dihapus", vbInformation, "Informasi"
    Call cmdBatal_Click
Exit Sub
errLoad:
    Call msubPesanError
    'MsgBox "Penghapusan data gagal", vbCritical, "Validasi"
End Sub

Private Function sp_Propinsi(f_status As String) As Boolean
    sp_Propinsi = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPropinsi", adChar, adParamInput, 2, txtKdPropinsi.Text)
        .Parameters.Append .CreateParameter("NamaPropinsi", adVarChar, adParamInput, 30, Trim(txtNamaPropinsi.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtPropinsi.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtPropinsi.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsPropinsi.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_Propinsi"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Master Propinsi", vbCritical, "Validasi"
            sp_Propinsi = False
        Else
            Call Add_HistoryLoginActivity("AUD_Propinsi")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_KotaKabupaten(f_status As String) As Boolean
sp_KotaKabupaten = True
    
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPropinsi", adChar, adParamInput, 2, dc2Propinsi.BoundText)
        .Parameters.Append .CreateParameter("KdKotaKabupaten", adVarChar, adParamInput, 4, txt2KdKotaKabupaten.Text)
        .Parameters.Append .CreateParameter("NamaKotaKabupaten", adVarChar, adParamInput, 50, Trim(txt2KotaKabupaten.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtKota.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtKota.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsKota.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_KotaKabupaten"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Master Kota/Kabupaten", vbCritical, "Validasi"
            sp_KotaKabupaten = False
        Else
            Call Add_HistoryLoginActivity("AUD_KotaKabupaten")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Kecamatan(f_status As String) As Boolean
sp_Kecamatan = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPropinsi", adChar, adParamInput, 2, dc3Propinsi.BoundText)
        .Parameters.Append .CreateParameter("KdKotaKabupaten", adVarChar, adParamInput, 4, dc3KotaKabupaten.BoundText)
        .Parameters.Append .CreateParameter("KdKecamatan", adVarChar, adParamInput, 6, txt3KdKecamatan.Text)
        .Parameters.Append .CreateParameter("NamaKecamatan", adVarChar, adParamInput, 50, Trim(txt3Kecamatan.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtKecamatan.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtKecamatan.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsKecamatan.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_Kecamatan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Master Kecamatan", vbCritical, "Validasi"
            sp_Kecamatan = False
        Else
            Call Add_HistoryLoginActivity("AUD_Kecamatan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Function sp_Kelurahan(f_status As String) As Boolean
    sp_Kelurahan = True
    Set adoCommand = New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPropinsi", adChar, adParamInput, 2, dc4Propinsi.BoundText)
        .Parameters.Append .CreateParameter("KdKotaKabupaten", adVarChar, adParamInput, 4, dc4KotaKabupaten.BoundText)
        .Parameters.Append .CreateParameter("KdKecamatan", adVarChar, adParamInput, 6, dc4Kecamatan.BoundText)
        .Parameters.Append .CreateParameter("KdKelurahan", adVarChar, adParamInput, 9, txt4KdKelurahanDesa.Text)
        .Parameters.Append .CreateParameter("KodePos", adVarChar, adParamInput, 10, Trim(txt4KodePos.Text))
        .Parameters.Append .CreateParameter("NamaKelurahan", adVarChar, adParamInput, 50, Trim(txt4KelurahanDesa.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExt.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExt.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_Kelurahan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan Master Kelurahan/Desa", vbCritical, "Validasi"
            sp_Kelurahan = False
        Else
            Call Add_HistoryLoginActivity("AUD_Kelurahan")
        End If
        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing
    End With
End Function

Private Sub cmdSimpan_Click()
On Error GoTo errLoad

    Select Case sstWilayah.Tab
        Case 0 'Propinsi
            If Periksa("text", txtNamaPropinsi, "Nama propinsi kosong") = False Then Exit Sub
            If sp_Propinsi("A") = False Then Exit Sub
        Case 1 'KotaKabupaten
            If Periksa("datacombo", dc2Propinsi, "Nama propinsi kosong") = False Then Exit Sub
            If Periksa("text", txt2KotaKabupaten, "Nama kota/kabupaten kosong") = False Then Exit Sub
            If sp_KotaKabupaten("A") = False Then Exit Sub
        Case 2 'Kecamatan
            If Periksa("datacombo", dc3Propinsi, "Nama propinsi kosong") = False Then Exit Sub
            If Periksa("datacombo", dc3KotaKabupaten, "Nama Kota/Kabupaten kosong") = False Then Exit Sub
            If Periksa("text", txt3Kecamatan, "Nama Kecamatan kosong") = False Then Exit Sub
            If sp_Kecamatan("A") = False Then Exit Sub
        Case 3 'Kelurahan/Desa
            If Periksa("datacombo", dc4Propinsi, "Nama propinsi kosong") = False Then Exit Sub
            If Periksa("datacombo", dc4KotaKabupaten, "Nama kota/kabupaten kosong") = False Then Exit Sub
            If Periksa("datacombo", dc4Kecamatan, "Nama Kecamatan kosong") = False Then Exit Sub
            If Periksa("text", txt4KodePos, "Kode pos kosong") = False Then Exit Sub
            If Periksa("text", txt4KelurahanDesa, "Nama kelurahan/desa kosong") = False Then Exit Sub
            If sp_Kelurahan("A") = False Then Exit Sub
    End Select
        
    MsgBox "Data berhasil disimpan", vbInformation, "Informasi"
    Call cmdBatal_Click
    
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dc2Propinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt2KotaKabupaten.SetFocus
End Sub

Private Sub dc3KotaKabupaten_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt3Kecamatan.SetFocus
End Sub

Private Sub dc3Propinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dc3KotaKabupaten.SetFocus
End Sub

Private Sub dc4Kecamatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt4KodePos.SetFocus
End Sub

Private Sub dc4KotaKabupaten_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dc4Kecamatan.SetFocus
End Sub

Private Sub dc4Propinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dc4KotaKabupaten.SetFocus
End Sub

Private Sub dgKecamatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt3Kecamatan.SetFocus
End Sub

Private Sub dgKelurahanDesa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt4KelurahanDesa.SetFocus
End Sub

Private Sub dgKotaKabupaten_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt2KotaKabupaten.SetFocus
End Sub

Private Sub dgPropinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaPropinsi.SetFocus
End Sub

Private Sub dgPropinsi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgPropinsi.ApproxCount = 0 Then Exit Sub
    txtKdPropinsi.Text = dgPropinsi.Columns(0).Value
    txtNamaPropinsi.Text = dgPropinsi.Columns(1).Value
    
    txtKdExtPropinsi.Text = dgPropinsi.Columns(2)
    txtNmExtPropinsi.Text = dgPropinsi.Columns(3)
    
    If dgPropinsi.Columns(4).Value = "<Type mismacth>" Then
        chkStsPropinsi.Value = 0
    Else
        If dgPropinsi.Columns(4).Value = 1 Then
            chkStsPropinsi.Value = 1
        Else
            chkStsPropinsi.Value = 0
        End If
    End If
    
End Sub

Private Sub dgKotaKabupaten_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgKotaKabupaten.ApproxCount = 0 Then Exit Sub
    dc2Propinsi.BoundText = dgKotaKabupaten.Columns(1).Value
    txt2KdKotaKabupaten.Text = dgKotaKabupaten.Columns(2).Value
    txt2KotaKabupaten.Text = dgKotaKabupaten.Columns(3).Value
    
    txtKdExtKota.Text = dgKotaKabupaten.Columns(4)
    txtNmExtKota.Text = dgKotaKabupaten.Columns(5)
    
    If dgKotaKabupaten.Columns(6).Value = "<Type mismacth>" Then
        chkStsKota.Value = 0
    Else
        If dgKotaKabupaten.Columns(6).Value = 1 Then
            chkStsKota.Value = 1
        Else
            chkStsKota.Value = 0
        End If
    End If
End Sub

Private Sub dgKecamatan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgKecamatan.ApproxCount = 0 Then Exit Sub
    dc3Propinsi.BoundText = dgKecamatan.Columns(1).Value
    dc3KotaKabupaten.BoundText = dgKecamatan.Columns(2).Value
    txt3KdKecamatan.Text = dgKecamatan.Columns(4).Value
    txt3Kecamatan.Text = dgKecamatan.Columns(5).Value
    
    txtKdExtKota.Text = dgKecamatan.Columns(6)
    txtNmExtKota.Text = dgKecamatan.Columns(7)
    
    If dgKecamatan.Columns(8).Value = "<Type mismacth>" Then
        chkStsKota.Value = 0
    Else
        If dgKecamatan.Columns(8).Value = 1 Then
            chkStsKota.Value = 1
        Else
            chkStsKota.Value = 0
        End If
    End If
End Sub

Private Sub dgKelurahanDesa_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgKelurahanDesa.ApproxCount = 0 Then Exit Sub
    dc4Propinsi.BoundText = dgKelurahanDesa.Columns(1).Value
    dc4KotaKabupaten.BoundText = dgKelurahanDesa.Columns(2).Value
    dc4Kecamatan.BoundText = dgKelurahanDesa.Columns(4).Value
    txt4KdKelurahanDesa.Text = dgKelurahanDesa.Columns(6).Value
    txt4KodePos.Text = dgKelurahanDesa.Columns(7).Value
    txt4KelurahanDesa.Text = dgKelurahanDesa.Columns(8).Value
    
    txtKdExt.Text = dgKelurahanDesa.Columns(9)
    txtNmExt.Text = dgKelurahanDesa.Columns(10)
    
    If dgKelurahanDesa.Columns(11).Value = "<Type mismacth>" Then
        chkSts.Value = 0
    Else
        If dgKelurahanDesa.Columns(11).Value = 1 Then
            chkSts.Value = 1
        Else
            chkSts.Value = 0
        End If
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo errLoad
    Call PlayFlashMovie(Me)
    Call centerForm(Me, MDIUtama)
    sstWilayah.Tab = 0
    Call cmdBatal_Click
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub clear()
    Select Case sstWilayah.Tab
        Case 0 'Propinsi
            txtKdPropinsi.Text = ""
            txtNamaPropinsi.Text = ""
            txtKdExtPropinsi = ""
            txtNmExtPropinsi = ""
        Case 1 'KotaKabupaten
            dc2Propinsi.BoundText = ""
            txt2KdKotaKabupaten.Text = ""
            txt2KotaKabupaten.Text = ""
            txtKdExtKota = ""
            txtNmExtKota = ""

        Case 2 'Kecamatan
            dc3Propinsi.BoundText = ""
            dc3KotaKabupaten.BoundText = ""
            txt3KdKecamatan.Text = ""
            txt3Kecamatan.Text = ""
            txtKdExtKecamatan = ""
            txtNmExtKecamatan = ""
        Case 3 'KelurahanDesa
            dc4Propinsi.BoundText = ""
            dc4KotaKabupaten.BoundText = ""
            dc4Kecamatan.BoundText = ""
            txt4KdKelurahanDesa.Text = ""
            txt4KodePos.Text = ""
            txt4KelurahanDesa.Text = ""
            txtKdExt = ""
            txtNmExt = ""
    End Select
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad
    Select Case sstWilayah.Tab
        Case 0
        Case 1
            Call msubDcSource(dc2Propinsi, rs, "Select * from Propinsi Where StatusEnabled = 1 ORDER BY NamaPropinsi")
        Case 2
            Call msubDcSource(dc3Propinsi, rs, "Select * from Propinsi Where StatusEnabled = 1  ORDER BY NamaPropinsi")
            Call msubDcSource(dc3KotaKabupaten, rs, "Select  KdKotaKabupaten, NamaKotaKabupaten from KotaKabupaten Where StatusEnabled = 1  ORDER BY NamaKotaKabupaten")
        Case 3
            Call msubDcSource(dc4Propinsi, rs, "Select * from Propinsi Where StatusEnabled = 1  ORDER BY NamaPropinsi")
            Call msubDcSource(dc4KotaKabupaten, rs, "Select  KdKotaKabupaten, NamaKotaKabupaten from KotaKabupaten Where StatusEnabled = 1  ORDER BY NamaKotaKabupaten")
            Call msubDcSource(dc4Kecamatan, rs, "Select   KdKecamatan, NamaKecamatan from Kecamatan Where StatusEnabled = 1  ORDER BY NamaKecamatan ")
    End Select
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub sstWilayah_Click(PreviousTab As Integer)
    Call clear
    Call subLoadDcSource
    Call subLoadGridSource
End Sub

Private Sub sstWilayah_DblClick()
    Call sstWilayah_KeyPress(13)
End Sub

Private Sub sstWilayah_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case sstWilayah.Tab
            Case 0 'Propinsi
                txtNamaPropinsi.SetFocus
            Case 1 'KotaKabupaten
                dc2Propinsi.SetFocus
            Case 2 'Kecamatan
                dc3Propinsi.SetFocus
            Case 3 'KelurahanDesa
                dc4Propinsi.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txt2KotaKabupaten_LostFocus()
    txt2KotaKabupaten.Text = Trim(StrConv(txt2KotaKabupaten.Text, vbProperCase))
End Sub

Private Sub txt3Kecamatan_LostFocus()
    txt3Kecamatan.Text = Trim(StrConv(txt3Kecamatan.Text, vbProperCase))
End Sub

Private Sub txt4KelurahanDesa_LostFocus()
    txt4KelurahanDesa.Text = Trim(StrConv(txt4KelurahanDesa.Text, vbProperCase))
End Sub

Private Sub txtNamaPropinsi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgPropinsi.SetFocus
    End Select
End Sub

Private Sub txtNamaPropinsi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txt2KotaKabupaten_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKotaKabupaten.SetFocus
    End Select
End Sub

Private Sub txt2KotaKabupaten_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txt3Kecamatan_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKecamatan.SetFocus
    End Select
End Sub

Private Sub txt3Kecamatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txt4KelurahanDesa_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKelurahanDesa.SetFocus
    End Select
End Sub

Private Sub txt4KelurahanDesa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txt4KodePos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txt4KelurahanDesa.SetFocus
End Sub

Private Sub txtNamaPropinsi_LostFocus()
    txtNamaPropinsi.Text = Trim(StrConv(txtNamaPropinsi.Text, vbProperCase))
End Sub
