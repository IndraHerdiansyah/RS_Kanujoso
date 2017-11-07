VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKomponenGaji 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Komponen Gaji"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7860
   Icon            =   "frmKomponenGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   7860
   Begin VB.Frame FrmKOmpoenGaji 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7815
      Begin VB.CommandButton cmdBatal 
         Caption         =   "Batal"
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
         Left            =   2040
         TabIndex        =   28
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdHapus 
         Caption         =   "Hapus"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "Simpan"
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
         Left            =   4920
         TabIndex        =   26
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutup"
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
         Left            =   6360
         TabIndex        =   25
         Top             =   6120
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   10186
         _Version        =   393216
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
         TabCaption(0)   =   "Komponen Gaji"
         TabPicture(0)   =   "frmKomponenGaji.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Komponen Potongan Gaji"
         TabPicture(1)   =   "frmKomponenGaji.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Mapping Jabatan"
         TabPicture(2)   =   "frmKomponenGaji.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(1)=   "chkStatus"
         Tab(2).Control(2)=   "dgKomponen"
         Tab(2).Control(3)=   "txtJumlah"
         Tab(2).Control(4)=   "dcJabatan"
         Tab(2).Control(5)=   "dckomponen"
         Tab(2).Control(6)=   "txtMasaKerja"
         Tab(2).Control(7)=   "lvKomponen"
         Tab(2).Control(8)=   "dcpendidikan"
         Tab(2).Control(9)=   "dcgolongan"
         Tab(2).Control(10)=   "dcKategory"
         Tab(2).Control(11)=   "Label17"
         Tab(2).Control(12)=   "Label16"
         Tab(2).Control(13)=   "Label15"
         Tab(2).Control(14)=   "Label14"
         Tab(2).Control(15)=   "Label13"
         Tab(2).Control(16)=   "Label12"
         Tab(2).Control(17)=   "Label11"
         Tab(2).Control(18)=   "Label10"
         Tab(2).Control(19)=   "Label9"
         Tab(2).ControlCount=   20
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   615
            Left            =   -71400
            TabIndex        =   47
            Top             =   720
            Width           =   1695
            Begin VB.OptionButton Option2 
               Caption         =   "Potongan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   49
               Top             =   240
               Width           =   1935
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Pendapatan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Value           =   -1  'True
               Width           =   1935
            End
         End
         Begin VB.CheckBox chkStatus 
            Caption         =   "Status Aktif"
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
            Left            =   -69120
            TabIndex        =   46
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1395
         End
         Begin MSDataGridLib.DataGrid dgKomponen 
            Height          =   2295
            Left            =   -74760
            TabIndex        =   45
            Top             =   3240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   -1  'True
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.TextBox txtJumlah 
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
            Left            =   -71400
            TabIndex        =   43
            Top             =   2760
            Width           =   2175
         End
         Begin MSDataListLib.DataCombo dcJabatan 
            Height          =   330
            Left            =   -74640
            TabIndex        =   29
            Top             =   1920
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dckomponen 
            Height          =   330
            Left            =   -71400
            TabIndex        =   41
            Top             =   1560
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   ""
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
         Begin VB.TextBox txtMasaKerja 
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
            Left            =   -71400
            TabIndex        =   40
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   5295
            Left            =   -74880
            TabIndex        =   12
            Top             =   360
            Width           =   7335
            Begin VB.CheckBox ChkKompPotonganGaji 
               Caption         =   "Status Aktif"
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
               Left            =   5640
               TabIndex        =   22
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1395
            End
            Begin VB.TextBox KdKompPotonganGaji 
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
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtNamaKompPotonganGaji 
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
               Height          =   375
               Left            =   2040
               TabIndex        =   15
               Top             =   480
               Width           =   5055
            End
            Begin VB.TextBox KdExtKompPotonganGaji 
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
               Height          =   375
               Left            =   120
               TabIndex        =   14
               Top             =   1200
               Width           =   1695
            End
            Begin VB.TextBox txtNamaExtKompPotonganGaji 
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
               Height          =   375
               Left            =   2040
               TabIndex        =   13
               Top             =   1200
               Width           =   3255
            End
            Begin MSDataGridLib.DataGrid dgKompPotonganGaji 
               Height          =   3495
               Left            =   120
               TabIndex        =   24
               Top             =   1680
               Width           =   7095
               _ExtentX        =   12515
               _ExtentY        =   6165
               _Version        =   393216
               AllowUpdate     =   -1  'True
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
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label8 
               Caption         =   "Kode Komponen"
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
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label7 
               Caption         =   "Nama Komponen Potongan Gaji"
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
               Left            =   2040
               TabIndex        =   19
               Top             =   240
               Width           =   2895
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label5 
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
               Height          =   255
               Left            =   2040
               TabIndex        =   17
               Top             =   960
               Width           =   1695
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Height          =   5295
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   7335
            Begin MSDataGridLib.DataGrid dgKompGaji 
               Height          =   3495
               Left            =   120
               TabIndex        =   23
               Top             =   1680
               Width           =   7095
               _ExtentX        =   12515
               _ExtentY        =   6165
               _Version        =   393216
               AllowUpdate     =   -1  'True
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
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.CheckBox ChkKompGaji 
               Caption         =   "Status Aktif"
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
               Left            =   5640
               TabIndex        =   21
               Top             =   1200
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtNamaExtKompGaji 
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
               Height          =   375
               Left            =   2040
               TabIndex        =   7
               Top             =   1200
               Width           =   3255
            End
            Begin VB.TextBox KdExtKompGaji 
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
               Height          =   375
               Left            =   120
               TabIndex        =   6
               Top             =   1200
               Width           =   1695
            End
            Begin VB.TextBox txtNamaKompGaji 
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
               Height          =   375
               Left            =   2040
               TabIndex        =   5
               Top             =   480
               Width           =   5055
            End
            Begin VB.TextBox KdKomponenGaji 
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
               Height          =   375
               Left            =   120
               TabIndex        =   4
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label4 
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
               Height          =   255
               Left            =   2040
               TabIndex        =   11
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label3 
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
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label2 
               Caption         =   "Nama Komponen Gaji"
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
               Left            =   2040
               TabIndex        =   9
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label1 
               Caption         =   "Kode Komponen"
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
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   1815
            End
         End
         Begin MSComctlLib.ListView lvKomponen 
            Height          =   255
            Left            =   -74760
            TabIndex        =   33
            Top             =   5280
            Visible         =   0   'False
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   450
            View            =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin MSDataListLib.DataCombo dcpendidikan 
            Height          =   330
            Left            =   -74640
            TabIndex        =   34
            Top             =   2520
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dcgolongan 
            Height          =   330
            Left            =   -74640
            TabIndex        =   35
            Top             =   1320
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   ""
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
         Begin MSDataListLib.DataCombo dcKategory 
            Height          =   330
            Left            =   -74640
            TabIndex        =   50
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   0
            Text            =   ""
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
         Begin VB.Label Label17 
            Caption         =   "Jumlah :"
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
            Left            =   -71400
            TabIndex        =   44
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Label Label16 
            Caption         =   "Thn"
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
            Left            =   -69480
            TabIndex        =   42
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label Label15 
            Caption         =   "Masa Kerja :"
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
            Left            =   -71400
            TabIndex        =   39
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Pendidikan :"
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
            Left            =   -74640
            TabIndex        =   38
            Top             =   2280
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Golongan :"
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
            Left            =   -74640
            TabIndex        =   37
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Jabatan :"
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
            Left            =   -74640
            TabIndex        =   36
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "Jenis         :"
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
            Left            =   -71400
            TabIndex        =   32
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Komponen :"
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
            Left            =   -71400
            TabIndex        =   31
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label9 
            Caption         =   "Type Pegawai :"
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
            Left            =   -74640
            TabIndex        =   30
            Top             =   480
            Width           =   1695
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
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
      Left            =   6000
      Picture         =   "frmKomponenGaji.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKomponenGaji.frx":1AA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
End
Attribute VB_Name = "frmKomponenGaji"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dcJabatan_Change()
'    If Option1.Value = True Then Call loadListViewSource_KomponenGaji
'    If Option2.Value = True Then Call loadListViewSource_KomponenPotGaji
End Sub

Private Sub dcJabatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        strSQL = "select * from jabatan where namajabatan like '%" & dcJabatan.Text & "%'"
        Call msubRecFO(rs, strSQL)
        If rs.RecordCount <> 0 Then
            dcJabatan.BoundText = rs!kdJabatan
            If Option1.Value = True Then Call loadListViewSource_KomponenGaji
            If Option2.Value = True Then Call loadListViewSource_KomponenPotGaji
        End If
    End If
End Sub

Private Sub dgKomponen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKomponen
        If .Columns("Jenis") = "Pendapatan" Then Option1.Value = True
        If .Columns("Jenis") = "Potongan" Then Option2.Value = True
        dcKategory.BoundText = .Columns("KdKategoryPegawai")
        dcpendidikan.BoundText = .Columns("KdPendidikan")
        dcJabatan.BoundText = .Columns("KdJabatan")
        dcGolongan.BoundText = .Columns("KdGolongan")
        dcKomponen.BoundText = .Columns("KdKomponen")
        txtMasaKErja.Text = .Columns("MasaKerja")
        txtJumlah.Text = .Columns("Jumlah")
        If .Columns("StatusEnabled") = "0" Then chkStatus.Value = Unchecked
        If .Columns("StatusEnabled") = "1" Then chkStatus.Value = Checked
    End With
End Sub

Private Sub Form_Load()
    On Error GoTo errLoad
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subLoadDcSource
    Call subLoadGrid
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub subLoadDcSource()
On Error GoTo errLoad

    Call msubDcSource(dcJabatan, rs, "select * from Jabatan where StatusEnabled ='1' order by KdJabatan")
    Call msubDcSource(dcpendidikan, rs, "select * from Pendidikan where StatusEnabled='1' order by NoUrut ")
    Call msubDcSource(dcGolongan, rs, "select * from GolonganPegawai where StatusEnabled='1' order by NoUrut ")
    Call msubDcSource(dcKategory, rs, "select * from TypePegawai  where StatusEnabled='1' ")
    Call loadListViewSource_KomponenGaji
    Option1.Value = True
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Public Sub subLoadGrid()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            Set rs = Nothing
            strSQL = "SELECT * from KomponenGaji"
            Call msubRecFO(rs, strSQL)
            Set dgKompGaji.DataSource = rs
        Case 1
            Set rs = Nothing
            strSQL = "SELECT * from KomponenPotonganGaji"
            Call msubRecFO(rs, strSQL)
            Set dgKompPotonganGaji.DataSource = rs
        Case 2
            Set rs = Nothing
            strSQL = "select typePegawai,namagolongan,NamaJabatan ,Pendidikan,jenis,KomponenGaji,MasaKerja ,Jumlah,StatusEnabled,KdKategoryPegawai,KdGolongan,KdPendidikan,KdJabatan,KdKomponen    from V_MappingJabatanKomponenGajiPotGaji order by typePegawai ,NamaGolongan,NamaJabatan,Pendidikan ,Jenis "
            Call msubRecFO(rs, strSQL)
            Set dgKomponen.DataSource = rs
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then Call loadListViewSource_KomponenGaji
End Sub

Public Sub loadListViewSource_KomponenGaji()
On Error GoTo errLoad
    mstrFilter = ""
    strSQL = "select * from KomponenGaji where statusEnabled='1'  order by KdKomponenGaji "
    Call msubDcSource(dcKomponen, rs, strSQL)
'    Call msubRecFO(rs, strSQL)
'    lvKomponen.ListItems.clear
'    While Not rs.EOF
'        lvKomponen.ListItems.add , "A" & rs(0).Value, rs(1).Value
'        rs.MoveNext
'    Wend
'    strSQL = "select * from V_mapPenggajian where statusEnabled='1' and jenis='Pendapatan'  order by KdKomponenGaji "
'    Call msubRecFO(rs, strSQL)
'    While Not rs.EOF
'        If IsNull(rs!kdJabatan) = True Then
'            lvKomponen.ListItems("A" & rs(0)).Checked = False
'            lvKomponen.ListItems("A" & rs(0)).ForeColor = vbBlack
'            lvKomponen.ListItems("A" & rs(0)).Bold = False
'        Else
'            If rs!kdJabatan = dcJabatan.BoundText Then ' And rs!Jenis = "Pendapatan" Then
'                lvKomponen.ListItems("A" & rs(0)).Checked = True
'                lvKomponen.ListItems("A" & rs(0)).ForeColor = vbBlue
'                lvKomponen.ListItems("A" & rs(0)).Bold = True
'            End If
'        End If
'        rs.MoveNext
'    Wend
    
Exit Sub
errLoad:
    Call msubPesanError
'    Resume 0
End Sub


Public Sub loadListViewSource_KomponenPotGaji()
On Error GoTo errLoad
    mstrFilter = ""
    strSQL = "select * from KomponenPotonganGaji where StatusEnabled='1' order by KdKomponenPotonganGaji "
    Call msubDcSource(dcKomponen, rs, strSQL)
    
'    Call msubRecFO(rs, strSQL)
'    lvKomponen.ListItems.clear
'    While Not rs.EOF
'        lvKomponen.ListItems.add , "A" & rs(0).Value, rs(1).Value
'        rs.MoveNext
'    Wend
'    strSQL = "select * from V_MapPenggajianPot where StatusEnabled='1' and jenis='Potongan' order by KdKomponenPotonganGaji "
'    Call msubRecFO(rs, strSQL)
''    lvKomponen.ListItems.clear
'    While Not rs.EOF
'         If IsNull(rs!kdJabatan) = True Then
'            lvKomponen.ListItems("A" & rs(0)).Checked = False
'            lvKomponen.ListItems("A" & rs(0)).ForeColor = vbBlack
'            lvKomponen.ListItems("A" & rs(0)).Bold = False
'        Else
'            If rs!kdJabatan = dcJabatan.BoundText Then 'And rs!Jenis = "Potongan" Then
'                lvKomponen.ListItems("A" & rs(0)).Checked = True
'                lvKomponen.ListItems("A" & rs(0)).ForeColor = vbBlue
'                lvKomponen.ListItems("A" & rs(0)).Bold = True
'            End If
'        End If
'        rs.MoveNext
'    Wend
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then loadListViewSource_KomponenPotGaji
End Sub

Private Sub Option4_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call subLoadGrid
End Sub

Private Sub dgKompGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKompGaji
        KdKomponenGaji.Text = .Columns(0)
        txtNamaKompGaji.Text = .Columns(1)
        KdExtKompGaji.Text = .Columns(2)
        txtNamaExtKompGaji.Text = .Columns(3)
        ChkKompGaji.Value = .Columns(4)
    End With
End Sub

Private Sub dgKompPotonganGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With dgKompPotonganGaji
        KdKompPotonganGaji.Text = .Columns(0)
        txtNamaKompPotonganGaji.Text = .Columns(1)
        KdExtKompPotonganGaji.Text = .Columns(2)
        txtNamaExtKompPotonganGaji.Text = .Columns(3)
        ChkKompPotonganGaji.Value = .Columns(4)
    End With
End Sub

Public Sub blank()
    Select Case SSTab1.Tab
        Case 0
            KdKomponenGaji.Text = ""
            txtNamaKompGaji.Text = ""
            KdExtKompGaji.Text = ""
            txtNamaExtKompGaji.Text = ""
            ChkKompGaji.Value = 1
        Case 1
            KdKompPotonganGaji.Text = ""
            txtNamaKompPotonganGaji.Text = ""
            KdExtKompPotonganGaji = ""
            txtNamaExtKompPotonganGaji.Text = ""
            ChkKompPotonganGaji.Value = 1
        Case 2
            dcKategory.Text = ""
            dcGolongan.Text = ""
            dcJabatan.Text = ""
            dcpendidikan.Text = ""
            dcKomponen.Text = ""
            txtMasaKErja.Text = ""
            txtJumlah.Text = ""
            chkStatus.Value = Checked
    End Select
End Sub

Public Sub cmdBatal_Click()
    Call blank
End Sub

Private Sub cmdHapus_Click()
    On Error GoTo bawah
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtNamaKompGaji, "Silahkan Pilih Nama Komponen") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "delete KomponenGaji where KdKomponenGaji = '" & KdKomponenGaji & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 1
            If Periksa("text", txtNamaKompPotonganGaji, "Silahkan Pilih Nama Komponen Potongan") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "delete KomponenPotonganGaji where KdKomponenPotonganGaji = '" & KdKompPotonganGaji.Text & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
        Case 2
            If Periksa("text", dcKategory, "Silahkan Pilih Nama Komponen Potongan") = False Then Exit Sub
            Set rs = Nothing
            If Option1.Value = True Then
                strSQL = "delete MappingJabatanKomponenGajiPotGaji where KdKategoryPegawai ='" & dcKategory.BoundText & "' and " & _
                        "KdGolongan='" & dcGolongan.BoundText & "' and " & _
                        "KdJabatan='" & dcJabatan.BoundText & "' and " & _
                        "KdPendidikan='" & dcpendidikan.BoundText & "' and " & _
                        "Jenis='Pendapatan' and " & _
                        "KdKomponen='" & dcKomponen.BoundText & "' "
            Else
                strSQL = "delete MappingJabatanKomponenGajiPotGaji where KdKategoryPegawai ='" & dcKategory.BoundText & "' and " & _
                        "KdGolongan='" & dcGolongan.BoundText & "' and " & _
                        "KdJabatan='" & dcJabatan.BoundText & "' and " & _
                        "KdPendidikan='" & dcpendidikan.BoundText & "' and " & _
                        "Jenis='Potongan' and " & _
                        "KdKomponen='" & dcKomponen.BoundText & "' "
            End If
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
    End Select
    Call blank
    Call subLoadGrid
    MsgBox "Data berhasil dihapus ", vbInformation, "Informasi"
    Exit Sub
bawah:
    MsgBox "Data Yang Akan Dihapus Masih Digunakan", vbExclamation, "Informasi"
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtNamaKompGaji, "Silahkan isi Nama Komponen Gaji") = False Then Exit Sub
            If sp_KomponenGaji("A") = False Then Exit Sub
            Call blank
        Case 1
            If Periksa("text", txtNamaKompPotonganGaji, "Silahkan isi Nama Komponen Potongan Gaji ") = False Then Exit Sub
            If sp_KomponenPotonganGaji("A") = False Then Exit Sub
            Call blank
        Case 2
            If dcKategory.Text = "" Then MsgBox "Silahkan Pilih Kategory Pegawai", vbInformation, "Mapping Jabatan": Exit Sub
            If dcKomponen.Text = "" Then MsgBox "Silahkan Pilih Komponen", vbInformation, "Mapping Jabatan": Exit Sub
'            If txtMasaKerja.Text = "" Then MsgBox "Silahkan Pilih Masa kerja", vbInformation, "Mapping Jabatan": Exit Sub
            If txtJumlah.Text = "" Then MsgBox "Silahkan Pilih Jumlah", vbInformation, "Mapping Jabatan": Exit Sub
            Dim ii As Integer
            Dim Jenis As String
            
            If Option1.Value = True Then
                Jenis = "Pendapatan"
            Else
                Jenis = "Potongan"
            End If
'            For ii = 1 To lvKomponen.ListItems.Count
'                If lvKomponen.ListItems(ii).Checked = True Then
                    If sp_MappingJabatanToPenggajian(Jenis, "a", "A") = False Then Exit Sub
'                Else
'                    If sp_MappingJabatanToPenggajian(Jenis, Right(lvKomponen.ListItems(ii).key, 2), "D") = False Then Exit Sub
'                End If
'            Next
            
    End Select

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call subLoadGrid
    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Function sp_MappingJabatanToPenggajian(Jenis As String, KdKomponen As String, f_status As String) As Boolean

    sp_MappingJabatanToPenggajian = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKategoryPegawai", adChar, adParamInput, 2, dcKategory.BoundText)
        .Parameters.Append .CreateParameter("KdGolongan", adVarChar, adParamInput, 2, dcGolongan.BoundText)
        .Parameters.Append .CreateParameter("KdJabatan", adVarChar, adParamInput, 5, dcJabatan.BoundText)
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 2, dcpendidikan.BoundText)
        .Parameters.Append .CreateParameter("Jenis", adVarChar, adParamInput, 30, Jenis)
        .Parameters.Append .CreateParameter("KdKomponen", adChar, adParamInput, 2, dcKomponen.BoundText)
        .Parameters.Append .CreateParameter("MasaKerja", adInteger, adParamInput, , Val(txtMasaKErja.Text))
        .Parameters.Append .CreateParameter("Jumlah", adDouble, adParamInput, , CDbl(txtJumlah.Text))
        .Parameters.Append .CreateParameter("Statusenabled", adChar, adParamInput, 1, chkStatus.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        .ActiveConnection = dbConn
        .CommandText = "AUD_MapJabatanToKomponenPenggajian"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Daftar Layanan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AUD_MapJabatanToKomponenPenggajian")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing

    End With

End Function

Private Function sp_KomponenGaji(f_status As String) As Boolean

    sp_KomponenGaji = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKomponenGaji", adChar, adParamInput, 2, KdKomponenGaji.Text)
        .Parameters.Append .CreateParameter("KomponenGaji", adVarChar, adParamInput, 50, txtNamaKompGaji.Text)
        .Parameters.Append .CreateParameter("kodeExternal", adChar, adParamInput, 15, IIf(KdExtKompGaji.Text = "", Null, KdExtKompGaji.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adChar, adParamInput, 50, IIf(txtNamaExtKompGaji.Text = "", Null, txtNamaExtKompGaji.Text))
        .Parameters.Append .CreateParameter("statusEnabled", adTinyInt, adParamInput, , ChkKompGaji.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        .ActiveConnection = dbConn
        .CommandText = "AUD_KomponenGaji"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Daftar Layanan", vbCritical, "Validasi"
        Else
            Call Add_HistoryLoginActivity("AUD_KomponenGaji")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing

    End With

End Function

Private Function sp_KomponenPotonganGaji(f_status As String) As Boolean

    sp_KomponenPotonganGaji = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKomponenPotonganGaji", adChar, adParamInput, 2, KdKompPotonganGaji.Text)
        .Parameters.Append .CreateParameter("KomponenPotonganGaji", adVarChar, adParamInput, 50, txtNamaKompPotonganGaji.Text)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(KdExtKompPotonganGaji.Text = "", Null, KdExtKompPotonganGaji.Text))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, IIf(txtNamaExtKompPotonganGaji.Text = "", Null, txtNamaExtKompPotonganGaji.Text))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , ChkKompPotonganGaji.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KomponenPotonganGaji"
        .CommandType = adCmdStoredProc
        .Execute
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam pemasukan data Daftar Layanan", vbCritical, "Validasi"

        Else
            Call Add_HistoryLoginActivity("AUD_KomponenPotonganGaji")
        End If

        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing

    End With

End Function

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub txtJumlah_Change()
    If IsNumeric(txtJumlah.Text) = False Then txtJumlah.Text = ""
End Sub

Private Sub txtMasaKerja_Change()
    If IsNumeric(txtMasaKErja.Text) = False Then txtMasaKErja.Text = ""
End Sub
