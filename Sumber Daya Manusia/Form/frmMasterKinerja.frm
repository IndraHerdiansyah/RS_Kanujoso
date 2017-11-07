VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterKinerja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Kategory & Detail Kategory Pegawai"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8955
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterKinerja.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6930
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8745
      Begin TabDlg.SSTab SSTab1 
         Height          =   6570
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   11589
         _Version        =   393216
         Tabs            =   5
         Tab             =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Jenis Kinerja"
         TabPicture(0)   =   "frmMasterKinerja.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Materi Kinerja"
         TabPicture(1)   =   "frmMasterKinerja.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Komponen Materi Kinerja"
         TabPicture(2)   =   "frmMasterKinerja.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtKomponenMateriKinerja"
         Tab(2).Control(1)=   "txtKdKomponenMateriKinerja"
         Tab(2).Control(2)=   "dgKomponenMateriKinerja"
         Tab(2).Control(3)=   "Label1(2)"
         Tab(2).Control(4)=   "Label1(1)"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Sasaran Kinerja"
         TabPicture(3)   =   "frmMasterKinerja.frx":0D1E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtNamaSasaranKinerja"
         Tab(3).Control(1)=   "txtKdSasaranKinerja"
         Tab(3).Control(2)=   "dcJenisKinerja"
         Tab(3).Control(3)=   "Frame4"
         Tab(3).Control(4)=   "Label12"
         Tab(3).Control(5)=   "Label10"
         Tab(3).Control(6)=   "Label9"
         Tab(3).ControlCount=   7
         TabCaption(4)   =   "Ukuran Kinerja"
         TabPicture(4)   =   "frmMasterKinerja.frx":0D3A
         Tab(4).ControlEnabled=   -1  'True
         Tab(4).Control(0)=   "Frame5"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.Frame Frame5 
            Height          =   5655
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   8055
            Begin MSDataGridLib.DataGrid dgUkuranKinerja 
               Height          =   3015
               Left            =   120
               TabIndex        =   60
               Top             =   2520
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   5318
               _Version        =   393216
               HeadLines       =   1
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
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox txtBobot 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2040
               MaxLength       =   5
               TabIndex        =   59
               Top             =   2040
               Width           =   855
            End
            Begin VB.TextBox txtSatuan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2040
               MaxLength       =   5
               TabIndex        =   58
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtUkuranKinerja 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   57
               Top             =   1320
               Width           =   5655
            End
            Begin VB.TextBox txtNamaUkuranKinerja 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2040
               MaxLength       =   50
               TabIndex        =   56
               Top             =   960
               Width           =   5655
            End
            Begin VB.TextBox txtKdUkuranKinerja 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2040
               MaxLength       =   3
               TabIndex        =   54
               Top             =   240
               Width           =   840
            End
            Begin MSDataListLib.DataCombo dcNamaSasaranKinerja 
               Height          =   330
               Left            =   2040
               TabIndex        =   55
               Top             =   600
               Width           =   4200
               _ExtentX        =   7408
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
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
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Bobot"
               Height          =   210
               Left            =   120
               TabIndex        =   53
               Top             =   2040
               Width           =   495
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Satuan"
               Height          =   210
               Left            =   120
               TabIndex        =   52
               Top             =   1680
               Width           =   570
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Ukuran Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   51
               Top             =   1320
               Width           =   1170
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Nama Ukuran Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   1680
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Nama Sasaran Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   49
               Top             =   600
               Width           =   1725
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Kode Ukuran Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   48
               Top             =   240
               Width           =   1650
            End
         End
         Begin VB.TextBox txtNamaSasaranKinerja 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   -72720
            MaxLength       =   50
            TabIndex        =   41
            Top             =   1680
            Width           =   5655
         End
         Begin VB.TextBox txtKdSasaranKinerja 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   330
            Left            =   -72720
            MaxLength       =   3
            TabIndex        =   37
            Top             =   960
            Width           =   840
         End
         Begin VB.TextBox txtKomponenMateriKinerja 
            Height          =   435
            Left            =   -72120
            TabIndex        =   33
            Top             =   1200
            Width           =   4935
         End
         Begin VB.TextBox txtKdKomponenMateriKinerja 
            Enabled         =   0   'False
            Height          =   315
            Left            =   -72120
            TabIndex        =   32
            Top             =   840
            Width           =   2295
         End
         Begin MSDataGridLib.DataGrid dgKomponenMateriKinerja 
            Height          =   4575
            Left            =   -74880
            TabIndex        =   31
            Top             =   1800
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   8070
            _Version        =   393216
            HeadLines       =   1
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
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5850
            Left            =   -74880
            TabIndex        =   19
            Top             =   1125
            Width           =   8265
            Begin VB.TextBox txtKdMateriKinerja 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   21
               Top             =   240
               Width           =   840
            End
            Begin VB.TextBox txtMateriKinerja 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   20
               Top             =   1320
               Width           =   5655
            End
            Begin MSDataGridLib.DataGrid dgMateriKinerja 
               Height          =   3135
               Left            =   120
               TabIndex        =   22
               Top             =   1800
               Width           =   8055
               _ExtentX        =   14208
               _ExtentY        =   5530
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
            Begin MSDataListLib.DataCombo dcRuangan 
               Height          =   330
               Left            =   2160
               TabIndex        =   23
               Top             =   600
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
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
            Begin MSDataListLib.DataCombo dcJabatan 
               Height          =   330
               Left            =   2160
               TabIndex        =   24
               Top             =   960
               Width           =   4200
               _ExtentX        =   7408
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
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
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Kode Materi Kinerja"
               Height          =   210
               Left            =   480
               TabIndex        =   28
               Top             =   285
               Width           =   1575
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Jabatan"
               Height          =   210
               Left            =   480
               TabIndex        =   27
               Top             =   1020
               Width           =   630
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Ruangan"
               Height          =   210
               Left            =   480
               TabIndex        =   26
               Top             =   645
               Width           =   705
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Materi Kinerja"
               Height          =   210
               Left            =   480
               TabIndex        =   25
               Top             =   1380
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Left            =   -74880
            TabIndex        =   8
            Top             =   840
            Width           =   8175
            Begin VB.TextBox txtJenisKinerja 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   35
               Top             =   720
               Width           =   5055
            End
            Begin VB.TextBox txtKdJenisKinerja 
               Enabled         =   0   'False
               Height          =   435
               Left            =   2160
               TabIndex        =   34
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txtKdExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   12
               Top             =   1080
               Width           =   1815
            End
            Begin VB.TextBox txtNamaExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   11
               Top             =   1440
               Width           =   5055
            End
            Begin VB.CheckBox CheckStatusEnbl 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               Height          =   255
               Left            =   5880
               TabIndex        =   10
               Top             =   1080
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtCariJenisKinerja 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1680
               MaxLength       =   30
               TabIndex        =   9
               Top             =   5400
               Width           =   6255
            End
            Begin MSDataGridLib.DataGrid dgJenisKinerja 
               Height          =   3330
               Left            =   120
               TabIndex        =   13
               Top             =   1920
               Width           =   7920
               _ExtentX        =   13970
               _ExtentY        =   5874
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode Jenis Kinerja"
               Height          =   210
               Index           =   0
               Left            =   480
               TabIndex        =   18
               Top             =   360
               Width           =   1470
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Kinerja"
               Height          =   210
               Left            =   480
               TabIndex        =   17
               Top             =   735
               Width           =   990
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   480
               TabIndex        =   16
               Top             =   1080
               Width           =   1140
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nama external"
               Height          =   210
               Left            =   480
               TabIndex        =   15
               Top             =   1440
               Width           =   1170
            End
            Begin VB.Label Label5 
               Caption         =   "Cari Jenis Kinerja"
               Height          =   255
               Left            =   240
               TabIndex        =   14
               Top             =   5450
               Width           =   1455
            End
         End
         Begin MSDataListLib.DataCombo dcJenisKinerja 
            Height          =   330
            Left            =   -72720
            TabIndex        =   39
            Top             =   1320
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
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
         Begin VB.Frame Frame4 
            Height          =   5655
            Left            =   -74880
            TabIndex        =   42
            Top             =   720
            Width           =   8055
            Begin MSDataGridLib.DataGrid dgSasaranKinerja 
               Height          =   3975
               Left            =   0
               TabIndex        =   46
               Top             =   1560
               Width           =   7815
               _ExtentX        =   13785
               _ExtentY        =   7011
               _Version        =   393216
               HeadLines       =   1
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
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Nama Sasaran Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   45
               Top             =   1080
               Width           =   1725
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   44
               Top             =   720
               Width           =   990
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Kode Sasaran Kinerja"
               Height          =   210
               Left            =   120
               TabIndex        =   43
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nama Sasaran Kerja"
            Height          =   210
            Left            =   -74760
            TabIndex        =   40
            Top             =   1800
            Width           =   1590
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Kinerja"
            Height          =   210
            Left            =   -74760
            TabIndex        =   38
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Kode Sasaran Kinerja"
            Height          =   210
            Left            =   -74760
            TabIndex        =   36
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Komponen Materi Kinerja"
            Height          =   210
            Index           =   2
            Left            =   -74880
            TabIndex        =   30
            Top             =   1320
            Width           =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode Komponen Materi Kinerja"
            Height          =   210
            Index           =   1
            Left            =   -74880
            TabIndex        =   29
            Top             =   840
            Width           =   2520
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   1
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
   Begin VB.Image Image4 
      Height          =   945
      Left            =   7080
      Picture         =   "frmMasterKinerja.frx":0D56
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterKinerja.frx":1ADE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterKinerja.frx":313C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmMasterKinerja.frx":5AFD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmMasterKinerja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub blankfield()
    On Error Resume Next
    Select Case SSTab1.Tab
        Case 0
            txtKdJenisKinerja.Text = ""
            txtJenisKinerja.Text = ""
            txtKdExt.Text = ""
            txtNamaExt.Text = ""
            CheckStatusEnbl.Value = 1
        Case 1
            txtKdMateriKinerja.Text = ""
            dcRuangan = ""
            dcJabatan = ""
            txtKdMateriKinerja = ""
            txtMateriKinerja = ""
        Case 2
            txtKdKomponenMateriKinerja.Text = ""
            txtKomponenMateriKinerja.Text = ""
        Case 3
            txtKdSasaranKinerja.Text = ""
            dcJenisKinerja = ""
            txtNamaSasaranKinerja.Text = ""
        Case 4
            txtKdUkuranKinerja.Text = ""
            dcNamaSasaranKinerja = ""
            txtNamaUkuranKinerja.Text = ""
            txtUkuranKinerja.Text = ""
            txtSatuan.Text = ""
            txtBobot.Text = ""
    End Select
End Sub

Sub Dag()
    On Error GoTo errLoad
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT * FROM dbo.JenisKinerja WHERE NamaJenisKinerja LIKE '%" & txtCariJenisKinerja.Text & "%'" 'WHERE (StatusEnabled <> 0) OR (StatusEnabled IS NULL)"
            Call msubRecFO(rs, strSQL)
            Set dgJenisKinerja.DataSource = rs
            dgJenisKinerja.Columns(0).Width = 500
            dgJenisKinerja.Columns(1).Width = 3000
            dgJenisKinerja.Columns(4).Width = 1000
        Case 1
            strSQL = "SELECT m.KdMateriKinerja, r.NamaRuangan, j.NamaJabatan, m.NamaMateriKinerja FROM dbo.MateriKinerja m INNER JOIN dbo.Ruangan r ON m.KdRuangan=r.KdRuangan INNER JOIN dbo.Jabatan j ON m.KdJabatan=j.KdJabatan where m.NamaMateriKinerja LIKE '%" & txtCariJenisKinerja.Text & "%'"
            Call msubRecFO(rs, strSQL)
            Set dgMateriKinerja.DataSource = rs
            dgMateriKinerja.Columns(0).Width = 2500
            dgMateriKinerja.Columns(1).Width = 1000
            dgMateriKinerja.Columns(2).Width = 1500
            dgMateriKinerja.Columns(3).Width = 3000
            Call msubDcSource(dcRuangan, rs, "SELECT * FROM dbo.Ruangan where statusenabled='1'")
            Call msubDcSource(dcJabatan, rs, "SELECT * FROM dbo.Jabatan where statusenabled='1'")
        Case 2
            strSQL = "SELECT * FROM dbo.KomponenMateriKinerja WHERE KomponenMateriKinerja LIKE '%" & txtCariJenisKinerja.Text & "%'"
            Call msubRecFO(rs, strSQL)
            Set dgKomponenMateriKinerja.DataSource = rs
            dgKomponenMateriKinerja.Columns(0).Width = 1500
            dgKomponenMateriKinerja.Columns(1).Width = 6500
        Case 3
            strSQL = "SELECT s.KdSasaranKinerja, j.NamaJenisKinerja, s.NamaSasaranKerja FROM dbo.SasaranKinerja s INNER JOIN dbo.JenisKinerja j on s.KdJenisKinerja=j.KdJenisKinerja WHERE s.NamaSasaranKerja LIKE '%" & txtCariJenisKinerja.Text & "%'"
            Call msubRecFO(rs, strSQL)
            Set dgSasaranKinerja.DataSource = rs
            dgSasaranKinerja.Columns(0).Width = 1500
            dgSasaranKinerja.Columns(1).Width = 3000
            dgSasaranKinerja.Columns(2).Width = 3000
            Call msubDcSource(dcJenisKinerja, rs, "SELECT * FROM dbo.JenisKinerja where statusenabled='1'")
        Case 4
            strSQL = "SELECT u.KdUkuranKinerja, s.NamaSasaranKerja, u.NamaKinerja, u.UkuranKinerja, u.Satuan, u.Bobot FROM dbo.UkuranKinerja u INNER JOIN dbo.SasaranKinerja s on u.KdSasaranKinerja=s.KdSasaranKinerja WHERE u.NamaKinerja LIKE '%" & txtCariJenisKinerja.Text & "%'"
            Call msubRecFO(rs, strSQL)
            Set dgUkuranKinerja.DataSource = rs
            dgUkuranKinerja.Columns(0).Width = 1500
            dgUkuranKinerja.Columns(1).Width = 2000
            dgUkuranKinerja.Columns(2).Width = 3000
            dgUkuranKinerja.Columns(3).Width = 1500
            dgUkuranKinerja.Columns(4).Width = 1500
            dgUkuranKinerja.Columns(5).Width = 700
            Call msubDcSource(dcNamaSasaranKinerja, rs, "SELECT KdSasaranKinerja, NamaSasaranKerja FROM SasaranKinerja ")
    End Select
    Exit Sub
errLoad:
    Call msubPesanError
    Set rs = Nothing
End Sub

Private Function sp_JenisPegawai(f_status As String) As Boolean
    sp_JenisPegawai = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisKinerja", adChar, adParamInput, 3, txtKdJenisKinerja.Text)
        .Parameters.Append .CreateParameter("NamaJenisKInerja", adVarChar, adParamInput, 50, Trim(txtJenisKinerja.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , CheckStatusEnbl.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisKinerja"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_JenisPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_MateriKinerja(f_status As String) As Boolean
    sp_MateriKinerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("kdRuangan", adChar, adParamInput, 3, dcRuangan.BoundText)
        .Parameters.Append .CreateParameter("kdJabatan", adVarChar, adParamInput, 5, dcJabatan.BoundText)
        .Parameters.Append .CreateParameter("KdMateriKinerja", adChar, adParamInput, 3, txtKdMateriKinerja.Text)
        .Parameters.Append .CreateParameter("NamaMateriKinerja", adVarChar, adParamInput, 50, txtMateriKinerja.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_MateriKinerja "
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_MateriKinerja = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_KomponenMateriKinerja(f_status As String) As Boolean
    sp_KomponenMateriKinerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKomponenMateriKinerja", adChar, adParamInput, 3, txtKdKomponenMateriKinerja.Text)
        .Parameters.Append .CreateParameter("KomponenMateriKinerja", adVarChar, adParamInput, 50, txtKomponenMateriKinerja.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KomponenMateriKinerja"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_KomponenMateriKinerja = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_SasaranKinerja(f_status As String) As Boolean
    sp_SasaranKinerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdSasaranKinerja", adChar, adParamInput, 3, txtKdSasaranKinerja.Text)
        .Parameters.Append .CreateParameter("KdJenisKinerja", adChar, adParamInput, 3, dcJenisKinerja.BoundText)
        .Parameters.Append .CreateParameter("NamaSasaranKerja", adVarChar, adParamInput, 50, txtNamaSasaranKinerja.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_SasaranKinerja"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_SasaranKinerja = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_UkuranKinerja(f_status As String) As Boolean
    sp_UkuranKinerja = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdUkuranKinerja", adChar, adParamInput, 3, txtKdUkuranKinerja.Text)
        .Parameters.Append .CreateParameter("KdSasaranKinerja", adChar, adParamInput, 3, dcNamaSasaranKinerja.BoundText)
        .Parameters.Append .CreateParameter("NamaKinerja", adVarChar, adParamInput, 250, txtNamaUkuranKinerja.Text)
        .Parameters.Append .CreateParameter("UkuranKinerja", adInteger, adParamInput, 3, txtUkuranKinerja.Text)
        .Parameters.Append .CreateParameter("Satuan", adVarChar, adParamInput, 15, txtSatuan.Text)
        .Parameters.Append .CreateParameter("Bobot", adDouble, adParamInput, , txtBobot.Text)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_UkuranKinerja"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_UkuranKinerja = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdCancel_Click()
    Call blankfield
    Call Dag
    Call SSTab1_KeyPress(13)
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini?", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisKinerja, "Nama Jenis Kinerja") = False Then Exit Sub
            If sp_JenisPegawai("D") = False Then Exit Sub
            Call blankfield
        Case 1
            If Periksa("text", txtMateriKinerja, "materi kinerja masih kosong") = False Then Exit Sub
            If sp_MateriKinerja("D") = False Then Exit Sub
            Call blankfield
        Case 2
            If Periksa("text", txtKomponenMateriKinerja, "komponen kinerja masih kosong") = False Then Exit Sub
            If sp_KomponenMateriKinerja("D") = False Then Exit Sub
            Call blankfield
        Case 3
            If Periksa("text", txtNamaSasaranKinerja, "Sasaran kinerja masih kosong") = False Then Exit Sub
            If Periksa("datacombo", dcJenisKinerja, "Silahkan pilih jenis kinerja") = False Then Exit Sub
            If sp_SasaranKinerja("D") = False Then Exit Sub
            Call blankfield
        Case 4
            If Periksa("text", txtNamaUkuranKinerja, "Nama Ukuran Kinerja masih kosong") = False Then Exit Sub
            If Periksa("datacombo", dcNamaSasaranKinerja, "Silahkan pilih nama sasaran kinerja") = False Then Exit Sub
            If sp_UkuranKinerja("D") = False Then Exit Sub
            Call blankfield
    End Select

    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errLoad

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisKinerja, "Silahkan isi Jenis Kinerja ") = False Then Exit Sub
            If sp_JenisPegawai("A") = False Then Exit Sub
            Call blankfield
        Case 1
            If Periksa("text", txtMateriKinerja, "Silahkan isi nama materi kinerja ") = False Then Exit Sub
            If sp_MateriKinerja("A") = False Then Exit Sub
            Call blankfield
        Case 2
            If Periksa("text", txtKomponenMateriKinerja, "Silahkan isi komponen materi kinerja ") = False Then Exit Sub
            If sp_KomponenMateriKinerja("A") = False Then Exit Sub
            Call blankfield
        Case 3
            If Periksa("text", txtNamaSasaranKinerja, "Silahkan isi komponen materi kinerja ") = False Then Exit Sub
            If Periksa("datacombo", dcJenisKinerja, "Silahkan pilih jenis kinerja ") = False Then Exit Sub
            If sp_SasaranKinerja("A") = False Then Exit Sub
            Call blankfield
        Case 4
            If Periksa("text", txtNamaUkuranKinerja, "Silahkan isi Nama Ukuran kinerja ") = False Then Exit Sub
            If Periksa("datacombo", dcNamaSasaranKinerja, "Silahkan pilih Sasaran kinerja ") = False Then Exit Sub
            If sp_UkuranKinerja("A") = False Then Exit Sub
            Call blankfield
    End Select

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dgJenisKinerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgJenisKinerja.ApproxCount = 0 Then Exit Sub
    txtKdJenisKinerja.Text = dgJenisKinerja.Columns(0).Value
    txtJenisKinerja.Text = dgJenisKinerja.Columns(1)
    If IsNull(dgJenisKinerja.Columns(2)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgJenisKinerja.Columns(2)
    If IsNull(dgJenisKinerja.Columns(3)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgJenisKinerja.Columns(3)
    CheckStatusEnbl.Value = dgJenisKinerja.Columns(4).Value
End Sub

Private Sub dgKomponenMateriKinerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKomponenMateriKinerja.ApproxCount = 0 Then Exit Sub
    txtKdKomponenMateriKinerja.Text = dgKomponenMateriKinerja.Columns(0).Value
    txtKomponenMateriKinerja.Text = dgKomponenMateriKinerja.Columns(1)
End Sub

Private Sub dgMateriKinerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgMateriKinerja.ApproxCount = 0 Then Exit Sub
    txtKdMateriKinerja.Text = dgMateriKinerja.Columns(0).Value
    dcRuangan.Text = dgMateriKinerja.Columns(1)
    dcJabatan.Text = dgMateriKinerja.Columns(2)
    txtMateriKinerja.Text = dgMateriKinerja.Columns(3)
End Sub

Private Sub dgSasaranKinerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgSasaranKinerja.ApproxCount = 0 Then Exit Sub
    txtKdSasaranKinerja.Text = dgSasaranKinerja.Columns(0).Value
    dcJenisKinerja.Text = dgSasaranKinerja.Columns(1)
    txtNamaSasaranKinerja.Text = dgSasaranKinerja.Columns(2)
End Sub

Private Sub dgUkuranKinerja_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgUkuranKinerja.ApproxCount = 0 Then Exit Sub
    txtKdUkuranKinerja.Text = dgUkuranKinerja.Columns(0).Value
    dcNamaSasaranKinerja.Text = dgUkuranKinerja.Columns(1)
    txtNamaUkuranKinerja.Text = dgUkuranKinerja.Columns(2)
    txtUkuranKinerja.Text = dgUkuranKinerja.Columns(3)
    txtSatuan.Text = dgUkuranKinerja.Columns(4)
    txtBobot.Text = dgUkuranKinerja.Columns(5)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 1
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call cmdCancel_Click
    Call blankfield
    SSTab1.Tab = 0
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdCancel_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
            
            Case 1
        
        End Select
    End If
errLoad:
End Sub

Private Sub txtCariJenisKinerja_Change()
    Call cmdCancel_Click
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtJenisKinerja1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgJenisKinerja.SetFocus
    End Select
End Sub

Private Sub txtJenisKinerja1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtKdJenisKinerja1_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisKinerja.SetFocus
End Sub

Private Sub txtNamaExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNamaExtDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub
