VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMasterInsentif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Insentif Pegawai"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasterInsentif.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9750
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   2880
      TabIndex        =   39
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   4200
      TabIndex        =   31
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   5520
      TabIndex        =   30
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6840
      TabIndex        =   29
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   8160
      TabIndex        =   28
      Top             =   8040
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
      Height          =   6810
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   9540
      Begin TabDlg.SSTab SSTab1 
         Height          =   6450
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   11377
         _Version        =   393216
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   882
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Jenis Insentif"
         TabPicture(0)   =   "frmMasterInsentif.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame4"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Komponen Insentif"
         TabPicture(1)   =   "frmMasterInsentif.frx":0CE6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Detail Komponen Insentif"
         TabPicture(2)   =   "frmMasterInsentif.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Map Insentif Pegawai"
         TabPicture(3)   =   "frmMasterInsentif.frx":0D1E
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "lvDetail"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "dcDetailview"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "chkSmua"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         Begin VB.CheckBox chkSmua 
            Caption         =   "Check Semua"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   5880
            Width           =   1815
         End
         Begin VB.Frame Frame4 
            Height          =   5775
            Left            =   -74760
            TabIndex        =   22
            Top             =   540
            Width           =   8895
            Begin VB.CheckBox chkSts 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6240
               TabIndex        =   37
               Top             =   960
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtNamaExtJenis 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   34
               Top             =   1320
               Width           =   5415
            End
            Begin VB.TextBox txtKdExtJenis 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   15
               TabIndex        =   33
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtJenisInsentif 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   24
               Top             =   600
               Width           =   5415
            End
            Begin VB.TextBox txtKdJenisInsentif 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               MaxLength       =   5
               TabIndex        =   23
               Top             =   240
               Width           =   975
            End
            Begin MSDataGridLib.DataGrid dgJenisInsentif 
               Height          =   3570
               Left            =   255
               TabIndex        =   25
               Top             =   1920
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   6297
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
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   210
               Left            =   240
               TabIndex        =   36
               Top             =   1380
               Width           =   1170
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   240
               TabIndex        =   35
               Top             =   1020
               Width           =   1140
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Kode"
               Height          =   210
               Left            =   240
               TabIndex        =   27
               Top             =   285
               Width           =   420
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Insentif"
               Height          =   210
               Left            =   240
               TabIndex        =   26
               Top             =   600
               Width           =   1065
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
            Height          =   5775
            Left            =   -74760
            TabIndex        =   17
            Top             =   540
            Width           =   8880
            Begin VB.CheckBox chkStsKomponen 
               Alignment       =   1  'Right Justify
               Caption         =   "Status Aktif"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6480
               TabIndex        =   38
               Top             =   1320
               Value           =   1  'Checked
               Width           =   1335
            End
            Begin VB.TextBox txtNamaKomponen 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   30
               TabIndex        =   1
               Top             =   600
               Width           =   5895
            End
            Begin VB.TextBox txtKdKomponen 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1920
               MaxLength       =   5
               TabIndex        =   0
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtKdExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   3
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox txtNamaExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   30
               TabIndex        =   4
               Top             =   1680
               Width           =   5895
            End
            Begin MSDataGridLib.DataGrid dgKomponen 
               Height          =   3330
               Left            =   255
               TabIndex        =   5
               Top             =   2160
               Width           =   8400
               _ExtentX        =   14817
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
            Begin MSDataListLib.DataCombo dcJenisInsentif 
               Height          =   330
               Left            =   1920
               TabIndex        =   2
               Top             =   960
               Width           =   2640
               _ExtentX        =   4657
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Style           =   2
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
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Insentif"
               Height          =   210
               Left            =   240
               TabIndex        =   32
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode"
               Height          =   210
               Left            =   240
               TabIndex        =   21
               Top             =   285
               Width           =   420
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nama Komponen"
               Height          =   210
               Left            =   240
               TabIndex        =   20
               Top             =   600
               Width           =   1395
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   240
               TabIndex        =   19
               Top             =   1380
               Width           =   1140
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nama external"
               Height          =   210
               Left            =   240
               TabIndex        =   18
               Top             =   1740
               Width           =   1170
            End
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
            Height          =   5715
            Left            =   -74760
            TabIndex        =   9
            Top             =   540
            Width           =   8820
            Begin VB.TextBox txtScore 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   75
               TabIndex        =   40
               Top             =   1440
               Width           =   1320
            End
            Begin VB.TextBox txtKdDetail 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1920
               MaxLength       =   5
               TabIndex        =   11
               Top             =   720
               Width           =   1320
            End
            Begin VB.TextBox txtDetail 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   30
               TabIndex        =   10
               Top             =   1080
               Width           =   5760
            End
            Begin MSDataGridLib.DataGrid dgDetail 
               Height          =   3600
               Left            =   240
               TabIndex        =   12
               Top             =   1920
               Width           =   8400
               _ExtentX        =   14817
               _ExtentY        =   6350
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
            Begin MSDataListLib.DataCombo dcKomponen 
               Height          =   330
               Left            =   1920
               TabIndex        =   13
               Top             =   360
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   582
               _Version        =   393216
               Appearance      =   0
               Style           =   2
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
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Score"
               Height          =   210
               Left            =   240
               TabIndex        =   41
               Top             =   1425
               Width           =   465
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Kode Detail"
               Height          =   210
               Left            =   240
               TabIndex        =   16
               Top             =   765
               Width           =   930
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Detail Komponen"
               Height          =   210
               Left            =   240
               TabIndex        =   15
               Top             =   1100
               Width           =   1395
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Nama Komponen"
               Height          =   210
               Left            =   240
               TabIndex        =   14
               Top             =   360
               Width           =   1395
            End
         End
         Begin MSDataListLib.DataCombo dcDetailview 
            Height          =   330
            Left            =   240
            TabIndex        =   43
            Top             =   840
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   582
            _Version        =   393216
            MatchEntry      =   -1  'True
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
         Begin MSComctlLib.ListView lvDetail 
            Height          =   4455
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   7858
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   8
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
      Left            =   7920
      Picture         =   "frmMasterInsentif.frx":0D3A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterInsentif.frx":1AC2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9975
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterInsentif.frx":3120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmMasterInsentif.frx":5AE1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmMasterInsentif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub clear()
    On Error Resume Next

    Select Case SSTab1.Tab
        Case 0
            txtKdJenisInsentif.Text = ""
            txtJenisInsentif.Text = ""
            txtKdExtJenis.Text = ""
            txtNamaExtJenis.Text = ""
            chkSts.Value = 1
            cmdDel.Enabled = True
            txtJenisInsentif.SetFocus
        Case 1
            txtKdKomponen.Text = ""
            txtNamaKomponen.Text = ""
            dcJenisInsentif.BoundText = ""
            txtKdExt.Text = ""
            txtNamaExt.Text = ""
            chkStsKomponen.Value = 1
            cmdDel.Enabled = True
            txtNamaKomponen.SetFocus
        Case 2
            dcKomponen.BoundText = ""
            txtKdDetail.Text = ""
            txtDetail.Text = ""
            txtScore.Text = ""
            cmdDel.Enabled = True
            dcKomponen.SetFocus
        Case 3
            dcDetailview.BoundText = ""
            cmdDel.Enabled = False

    End Select
End Sub

Sub subLoadDcSource()
    Select Case SSTab1.Tab
        Case 0
        Case 1
            Call msubDcSource(dcJenisInsentif, rs, "SELECT KdJnsInsentif, Jenisinsentif FROM JenisInsentif where statusenabled='1' order by JenisInsentif")
        Case 2
            Call msubDcSource(dcKomponen, rs, "SELECT KdKomponenInsentif, KomponenInsentif FROM KomponenInsentif where statusenabled='1' order by KomponenInsentif")
        Case 3 'DetailKomponenInsentif
            strSQL = "select KdDetailKomponenInsentif,DetailKomponenInsentif from DetailKomponenInsentif order by DetailKomponenInsentif"
            Call msubDcSource(dcDetailview, rs, strSQL)
    End Select
End Sub

Sub subLoadGridSource()
    On Error GoTo hell
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT * FROM JenisInsentif order by JenisInsentif "
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgJenisInsentif.DataSource = rs
            With dgJenisInsentif
                .Columns(0).Width = 2000
                .Columns(0).Caption = "Kode"
                .Columns(1).Width = 4500
                .Columns(1).Caption = "Jenis Insentif"
            End With

        Case 1
            strSQL = "SELECT dbo.KomponenInsentif.KdKomponenInsentif AS Kode, dbo.KomponenInsentif.KomponenInsentif, dbo.JenisInsentif.JenisInsentif AS [Jenis Insentif], " & _
            "dbo.KomponenInsentif.KodeExternal AS [Kd.Ext], dbo.KomponenInsentif.NamaExternal AS [Nama Ext], dbo.KomponenInsentif.KdJnsInsentif, dbo.KomponenInsentif.StatusEnabled " & _
            "FROM dbo.KomponenInsentif LEFT OUTER JOIN " & _
            "dbo.JenisInsentif ON dbo.KomponenInsentif.KdJnsInsentif = dbo.JenisInsentif.KdJnsInsentif order by dbo.KomponenInsentif.KomponenInsentif "
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgKomponen.DataSource = rs
            With dgKomponen
                .Columns(6).Width = 0
            End With

        Case 2
            strSQL = "SELECT dbo.DetailKomponenInsentif.KdDetailKomponenInsentif AS Kode, dbo.DetailKomponenInsentif.DetailKomponenInsentif AS Detail, dbo.KomponenInsentif.KomponenInsentif, " & _
            "dbo.DetailKomponenInsentif.Score, " & _
            "dbo.DetailKomponenInsentif.KdKomponenInsentif " & _
            "FROM dbo.DetailKomponenInsentif LEFT OUTER JOIN " & _
            "dbo.KomponenInsentif ON dbo.DetailKomponenInsentif.KdKomponenInsentif = dbo.KomponenInsentif.KdKomponenInsentif"
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgDetail.DataSource = rs
            With dgDetail
                .Columns(1).Width = 2500
                .Columns("KdKomponenInsentif").Width = 0
            End With
        Case 3
            cmdDel.Enabled = False
            Call loadListViewSource
    End Select
    Exit Sub
hell:
    Call msubPesanError
End Sub

Private Function sp_JenisInsentif(f_Status As String) As Boolean
    sp_JenisInsentif = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJnsInsentif", adTinyInt, adParamInput, , IIf(txtKdJenisInsentif.Text = "", Null, Trim(txtKdJenisInsentif.Text)))
        .Parameters.Append .CreateParameter("JenisInsentif", adVarChar, adParamInput, 30, Trim(txtJenisInsentif.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExtJenis.Text = "", Null, Trim(txtKdExtJenis.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExtJenis.Text = "", Null, Trim(txtNamaExtJenis.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisInsentif"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_JenisInsentif = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_KomponenInsentif(f_Status As String) As Boolean
    sp_KomponenInsentif = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKomponenInsentif", adSmallInt, adParamInput, , IIf(txtKdKomponen.Text = "", Null, Trim(txtKdKomponen.Text)))
        .Parameters.Append .CreateParameter("KomponenInsentif", adVarChar, adParamInput, 30, Trim(txtNamaKomponen.Text))
        .Parameters.Append .CreateParameter("KdJnsInsentif", adTinyInt, adParamInput, , IIf(dcJenisInsentif.Text = "", Null, Trim(dcJenisInsentif.BoundText)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 30, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsKomponen.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_KomponenInsentif"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_KomponenInsentif = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_DetailKomponenInsentif(f_Status As String) As Boolean
    sp_DetailKomponenInsentif = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdDetailKomponenInsentif", adSmallInt, adParamInput, , IIf(txtKdDetail.Text = "", Null, Trim(txtKdDetail.Text)))
        .Parameters.Append .CreateParameter("DetailKomponenInsentif", adVarChar, adParamInput, 30, Trim(txtDetail.Text))
        .Parameters.Append .CreateParameter("KdKomponenInsentif", adSmallInt, adParamInput, , dcKomponen.BoundText)
        .Parameters.Append .CreateParameter("Score", adDouble, adParamInput, , IIf(txtScore.Text = "", 0, Trim(txtScore.Text)))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_DetailKomponenInsentif"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_DetailKomponenInsentif = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdCancel_Click()
    Call clear
    Call subLoadDcSource
    Call subLoadGridSource
    Call loadListViewSource
    Call SSTab1_KeyPress(13)
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisInsentif, "Jenis Insentif kosong") = False Then Exit Sub
            If sp_JenisInsentif("D") = False Then Exit Sub
        Case 1
            If Periksa("text", txtNamaKomponen, "Nama Komponen Insentif kosong ") = False Then Exit Sub
            If sp_KomponenInsentif("D") = False Then Exit Sub
        Case 2
            If Periksa("text", txtDetail, "Detail Komponen Insentif kosong ") = False Then Exit Sub
            If sp_DetailKomponenInsentif("D") = False Then Exit Sub
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
            If Periksa("text", txtJenisInsentif, "Silahkan isi Jenis Insentif ") = False Then Exit Sub
            If sp_JenisInsentif("A") = False Then Exit Sub

        Case 1
            If Periksa("text", txtNamaKomponen, "Silahkan isi Komponen Insentif ") = False Then Exit Sub
            If Periksa("datacombo", dcJenisInsentif, "Silahkan isi Jenis Insentif ") = False Then Exit Sub
            If sp_KomponenInsentif("A") = False Then Exit Sub

        Case 2
            If Periksa("text", txtDetail, "Silahkan isi Detail Komponen Insentif ") = False Then Exit Sub
            If Periksa("datacombo", dcKomponen, "Silahkan isi Komponen Insentif ") = False Then Exit Sub
            If sp_DetailKomponenInsentif("A") = False Then Exit Sub
        Case 3
            If Periksa("datacombo", dcDetailview, "Silahkan isi Detail komponen Insentif") = False Then Exit Sub

            For i = 1 To lvDetail.ListItems.Count
                If lvDetail.ListItems(i).Checked = True Then
                    If sp_MapInsentifPegawai(Right(lvDetail.ListItems(i).key, Len(lvDetail.ListItems(i).key) - 1), "A") = False Then Exit Sub
                Else
                    If sp_MapInsentifPegawai(Right(lvDetail.ListItems(i).key, Len(lvDetail.ListItems(i).key) - 1), "D") = False Then Exit Sub
                End If
            Next i
    End Select

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisInsentif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub dcKomponen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtScore.SetFocus
End Sub

Private Sub dcDetailview_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvDetail.SetFocus
End Sub

Private Sub dgJenisInsentif_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisInsentif
    WheelHook.WheelHook dgJenisInsentif
End Sub

Private Sub dgJenisInsentif_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgJenisInsentif.ApproxCount = 0 Then Exit Sub
    txtKdJenisInsentif.Text = dgJenisInsentif.Columns(0).Value
    txtJenisInsentif.Text = dgJenisInsentif.Columns(1).Value
    txtKdExtJenis.Text = dgJenisInsentif.Columns(2).Value
    txtNamaExtJenis.Text = dgJenisInsentif.Columns(3).Value
    If dgJenisInsentif.Columns(4).Value = "<Type mismacth>" Then
        chkSts.Value = 0
    Else
        If dgJenisInsentif.Columns(4).Value = 1 Then
            chkSts.Value = 1
        Else
            chkSts.Value = 0
        End If
    End If
End Sub

Private Sub dgDetail_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgDetail
    WheelHook.WheelHook dgDetail
End Sub

Private Sub dgDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDetail.SetFocus
End Sub

Private Sub dgDetail_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgDetail.ApproxCount = 0 Then Exit Sub
    dcKomponen.BoundText = dgDetail.Columns(4)
    txtKdDetail = dgDetail.Columns(0)
    txtDetail = dgDetail.Columns(1)
    txtScore = dgDetail.Columns(3)

End Sub

Private Sub dgKomponen_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgKomponen
    WheelHook.WheelHook dgKomponen
End Sub

Private Sub dgKomponen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaKomponen.SetFocus
End Sub

Private Sub dgKomponen_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgKomponen.ApproxCount = 0 Then Exit Sub
    txtKdKomponen.Text = dgKomponen.Columns(0)
    txtNamaKomponen.Text = dgKomponen.Columns(1)

    If IsNull(dgKomponen.Columns(5)) Then dcJenisInsentif.BoundText = "" Else dcJenisInsentif.BoundText = dgKomponen.Columns(5)
    If IsNull(dgKomponen.Columns(3)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgKomponen.Columns(3)
    If IsNull(dgKomponen.Columns(4)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgKomponen.Columns(4)
    chkStsKomponen.Value = dgKomponen.Columns("StatusEnabled").Value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCtrlKey As String
    strCtrlKey = (Shift + vbCtrlMask)
    Select Case KeyCode
        Case vbKey1
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 0
        Case vbKey2
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 1
        Case vbKey3
            If strCtrlKey = 4 Then SSTab1.SetFocus: SSTab1.Tab = 2
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    SSTab1.Tab = 0

    Call cmdCancel_Click

End Sub

Private Sub lvDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdCancel_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtNamaKomponen.SetFocus
            Case 1
                dcKomponen.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtJenisInsentif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtJenis.SetFocus
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkStsKomponen.SetFocus
End Sub

Private Sub txtKdExtJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkSts.SetFocus
End Sub

Private Sub txtKdKomponen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaKomponen.SetFocus
End Sub

Private Sub txtKdDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDetail.SetFocus
End Sub

Private Sub txtDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            txtScore.SetFocus
    End Select
End Sub

Private Sub txtDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtScore.SetFocus
End Sub

Private Sub chkSts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExtJenis.SetFocus
End Sub

Private Sub chkStsKomponen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtNamaExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNamaExtJenis_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNoUrut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisInsentif.SetFocus
End Sub

Private Sub txtNamaKomponen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisInsentif.SetFocus
End Sub

Private Sub txtNamaExtDetail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtRepDisplay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtScore_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub lvDetail_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If lvDetail.ListItems(Item.key).Checked = True Then
        lvDetail.ListItems(Item.key).ForeColor = vbBlue
    Else
        lvDetail.ListItems(Item.key).ForeColor = vbBlack
    End If

End Sub

Private Sub dcDetailview_Change()

    On Error Resume Next

    Call loadListViewSource

End Sub

Public Sub loadListViewSource()

    On Error GoTo errLoad

    strSQL = "select IdPegawai, NamaLengkap from DataPegawai order by NamaLengkap"
    Call msubRecFO(rs, strSQL)
    lvDetail.ListItems.clear

    While Not rs.EOF
        lvDetail.ListItems.add , "A" & rs(0).Value, rs(1).Value
        rs.MoveNext
    Wend

    lvDetail.Sorted = True

    strSQL = "select MapInsentifPegawai.IdPegawai from MapInsentifPegawai inner join DataPegawai " & _
    " on MapInsentifPegawai.IdPegawai = DataPegawai.IdPegawai " & _
    " where MapInsentifPegawai.KdDetailKomponenInsentif = '" & dcDetailview.BoundText & "'"

    Call msubRecFO(rs, strSQL)

    Do While rs.EOF = False
        lvDetail.ListItems("A" & rs(0)).Checked = True
        lvDetail.ListItems("A" & rs(0)).ForeColor = vbBlue
        lvDetail.ListItems("A" & rs(0)).Bold = True
        rs.MoveNext
    Loop

    Exit Sub

errLoad:
    Call msubPesanError

End Sub

Private Sub chkSmua_Click()

    If chkSmua.Value = Checked Then
        For i = 1 To lvDetail.ListItems.Count
            lvDetail.ListItems.Item(i).Checked = True
        Next i
    Else
        For i = 1 To lvDetail.ListItems.Count
            lvDetail.ListItems.Item(i).Checked = False
        Next i
    End If

End Sub

Private Function sp_MapInsentifPegawai(f_IdPegawai As String, f_Status As String) As Boolean

    On Error GoTo errLoad

    sp_MapInsentifPegawai = True

    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, f_IdPegawai)
        .Parameters.Append .CreateParameter("KdDetailKomponenInsentif", adSmallInt, adParamInput, , dcDetailview.BoundText)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AU_MapInsentifPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value") <> 0 Then
            MsgBox "Ada kesalahan dalam penyimpanan data ", vbCritical, "Validasi"
            sp_MapInsentifPegawai = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With

    Exit Function

errLoad:

End Function

