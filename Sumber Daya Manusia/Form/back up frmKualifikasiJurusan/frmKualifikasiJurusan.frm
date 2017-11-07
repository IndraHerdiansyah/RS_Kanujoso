VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmKualifikasiJurusan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Pendidikan dan Kualifikasi Jurusan"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmKualifikasiJurusan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   9015
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   3360
      TabIndex        =   41
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   4680
      TabIndex        =   40
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   6000
      TabIndex        =   39
      Top             =   8160
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   7320
      TabIndex        =   38
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
      TabIndex        =   9
      Top             =   1080
      Width           =   8820
      Begin TabDlg.SSTab SSTab1 
         Height          =   6450
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   11377
         _Version        =   393216
         Tab             =   1
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
         TabCaption(0)   =   "Jenis Pendidikan"
         TabPicture(0)   =   "frmKualifikasiJurusan.frx":0CCA
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Pendidikan"
         TabPicture(1)   =   "frmKualifikasiJurusan.frx":0CE6
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Kualifikasi Jurusan"
         TabPicture(2)   =   "frmKualifikasiJurusan.frx":0D02
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame Frame4 
            Height          =   5535
            Left            =   -74760
            TabIndex        =   32
            Top             =   600
            Width           =   7935
            Begin VB.TextBox txtJenisPendidikan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   2160
               MaxLength       =   30
               TabIndex        =   34
               Top             =   1080
               Width           =   5415
            End
            Begin VB.TextBox txtKdJenisPendidikan 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   33
               Top             =   720
               Width           =   975
            End
            Begin MSDataGridLib.DataGrid dgJenisPendidikan 
               Height          =   3690
               Left            =   255
               TabIndex        =   35
               Top             =   1560
               Width           =   7320
               _ExtentX        =   12912
               _ExtentY        =   6509
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
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Kode Jenis"
               Height          =   210
               Left            =   240
               TabIndex        =   37
               Top             =   765
               Width           =   870
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Nama Jenis Pendidikan"
               Height          =   210
               Left            =   240
               TabIndex        =   36
               Top             =   1080
               Width           =   1830
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
            Height          =   5655
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   8040
            Begin VB.TextBox txtPendidikan 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   1
               Top             =   600
               Width           =   5055
            End
            Begin VB.TextBox txtKdPendidikan 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1920
               MaxLength       =   4
               TabIndex        =   0
               Top             =   240
               Width           =   975
            End
            Begin VB.TextBox txtRepDisplay 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   25
               TabIndex        =   4
               Top             =   1800
               Width           =   5055
            End
            Begin VB.TextBox txtKdExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   5
               Top             =   2160
               Width           =   5055
            End
            Begin VB.TextBox txtNamaExt 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   30
               TabIndex        =   6
               Top             =   2520
               Width           =   5055
            End
            Begin VB.TextBox txtNoUrut 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   2
               TabIndex        =   2
               Top             =   960
               Width           =   975
            End
            Begin MSDataGridLib.DataGrid dgPendidikan 
               Height          =   2490
               Left            =   255
               TabIndex        =   7
               Top             =   3000
               Width           =   7560
               _ExtentX        =   13335
               _ExtentY        =   4392
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
            Begin MSDataListLib.DataCombo dcJenisPendidikan 
               Height          =   330
               Left            =   1920
               TabIndex        =   3
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
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Jenis Pendidikan"
               Height          =   210
               Left            =   240
               TabIndex        =   42
               Top             =   1320
               Width           =   1320
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Kode Pendidikan"
               Height          =   210
               Left            =   240
               TabIndex        =   31
               Top             =   285
               Width           =   1350
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Nama Pendidikan"
               Height          =   210
               Left            =   240
               TabIndex        =   30
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Report Display"
               Height          =   210
               Left            =   240
               TabIndex        =   29
               Top             =   1860
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   240
               TabIndex        =   28
               Top             =   2220
               Width           =   1140
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Nama external"
               Height          =   210
               Left            =   240
               TabIndex        =   27
               Top             =   2580
               Width           =   1170
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "No. Urut"
               Height          =   210
               Left            =   240
               TabIndex        =   26
               Top             =   960
               Width           =   705
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
            TabIndex        =   11
            Top             =   480
            Width           =   7980
            Begin VB.TextBox txtKdKualifikasi 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
               Height          =   330
               Left            =   1920
               MaxLength       =   4
               TabIndex        =   16
               Top             =   720
               Width           =   1320
            End
            Begin VB.TextBox txtKualifikasi 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   75
               TabIndex        =   15
               Top             =   1080
               Width           =   5760
            End
            Begin VB.TextBox txtRepDisplayKualifikasi 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   75
               TabIndex        =   14
               Top             =   1680
               Width           =   5760
            End
            Begin VB.TextBox txtKdExtKualifikasi 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   15
               TabIndex        =   13
               Top             =   2040
               Width           =   3840
            End
            Begin VB.TextBox txtNamaExtKualifikasi 
               Appearance      =   0  'Flat
               Height          =   330
               Left            =   1920
               MaxLength       =   30
               TabIndex        =   12
               Top             =   2400
               Width           =   3840
            End
            Begin MSDataGridLib.DataGrid dgKualifikasi 
               Height          =   2640
               Left            =   240
               TabIndex        =   17
               Top             =   2880
               Width           =   7440
               _ExtentX        =   13123
               _ExtentY        =   4657
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
            Begin MSDataListLib.DataCombo dcPendidikan 
               Height          =   330
               Left            =   1920
               TabIndex        =   18
               Top             =   360
               Width           =   3720
               _ExtentX        =   6562
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
               Caption         =   "Kode Kualifikasi"
               Height          =   210
               Left            =   240
               TabIndex        =   24
               Top             =   765
               Width           =   1215
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Kualifikasi Jurusan"
               Height          =   210
               Left            =   240
               TabIndex        =   23
               Top             =   1140
               Width           =   1410
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Nama Pendidikan"
               Height          =   210
               Left            =   240
               TabIndex        =   22
               Top             =   360
               Width           =   1380
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Report Display"
               Height          =   210
               Left            =   240
               TabIndex        =   21
               Top             =   1620
               Width           =   1155
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Kode External"
               Height          =   210
               Left            =   240
               TabIndex        =   20
               Top             =   1980
               Width           =   1140
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Nama External"
               Height          =   210
               Left            =   240
               TabIndex        =   19
               Top             =   2340
               Width           =   1170
            End
         End
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   10
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
      Left            =   7200
      Picture         =   "frmKualifikasiJurusan.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKualifikasiJurusan.frx":1AA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmKualifikasiJurusan.frx":3104
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmKualifikasiJurusan.frx":5AC5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmKualifikasiJurusan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub clear()
On Error Resume Next

Select Case SSTab1.Tab
    Case 0
        txtKdJenisPendidikan.Text = ""
        txtJenisPendidikan.Text = ""
        txtJenisPendidikan.SetFocus
     
     Case 1
        txtKdPendidikan.Text = ""
        txtPendidikan.Text = ""
        txtNoUrut.Text = ""
        dcJenisPendidikan.BoundText = ""
        txtRepDisplay.Text = ""
        txtKdExt.Text = ""
        txtNamaExt.Text = ""
        txtPendidikan.SetFocus
     
    Case 2
        dcPendidikan.BoundText = ""
        txtKdKualifikasi.Text = ""
        txtKualifikasi.Text = ""
        txtRepDisplayKualifikasi.Text = ""
        txtKdExtKualifikasi.Text = ""
        txtNamaExtKualifikasi.Text = ""
        dcPendidikan.SetFocus
End Select
End Sub

Sub SubLoadDCSource()
Select Case SSTab1.Tab
    Case 0
    Case 1
        Call msubDcSource(dcJenisPendidikan, rs, "SELECT KdJenisPendidikan, JenisPendidikan FROM JenisPendidikan order by JenisPendidikan")
    Case 2
        Call msubDcSource(dcPendidikan, rs, "SELECT KdPendidikan, Pendidikan FROM Pendidikan where StatusEnabled = '1' order by Pendidikan")
End Select
End Sub

Sub subLoadGridSource()
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT * FROM JenisPendidikan order by JenisPendidikan "
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgJenisPendidikan.DataSource = rs
            With dgJenisPendidikan
                .Columns(0).Width = 2000
                .Columns(0).Caption = "Kode Jenis"
                .Columns(1).Width = 4500
                .Columns(1).Caption = "Jenis Pendidikan"
                
            End With
        
        Case 1
            strSQL = "SELECT dbo.Pendidikan.KdPendidikan AS Kode, dbo.Pendidikan.Pendidikan, dbo.Pendidikan.NoUrut AS [No. Urut], dbo.JenisPendidikan.JenisPendidikan AS [Jenis Pendidikan], " & _
                     "dbo.Pendidikan.KodeExternal AS [Kd.Ext], dbo.Pendidikan.NamaExternal AS [Nama Ext], dbo.Pendidikan.KdJenisPendidikan, dbo.Pendidikan.StatusEnabled " & _
                     "FROM dbo.Pendidikan LEFT OUTER JOIN " & _
                     "dbo.JenisPendidikan ON dbo.Pendidikan.KdJenisPendidikan = dbo.JenisPendidikan.KdJenisPendidikan " & _
                     "WHERE (dbo.Pendidikan.StatusEnabled <> 0) OR (dbo.Pendidikan.StatusEnabled IS NULL) order by dbo.Pendidikan.NoUrut "
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgPendidikan.DataSource = rs
            With dgPendidikan
                .Columns(6).Width = 0
            End With
        
        Case 2
            strSQL = "SELECT dbo.KualifikasiJurusan.KdKualifikasiJurusan AS KODE, dbo.KualifikasiJurusan.KualifikasiJurusan AS Jurusan, dbo.Pendidikan.Pendidikan, " & _
                     "dbo.KualifikasiJurusan.ReportDisplay AS [Rep.Display], dbo.KualifikasiJurusan.KodeExternal AS [Kd.Ext], dbo.KualifikasiJurusan.NamaExternal AS [Nama Ext], " & _
                     "dbo.KualifikasiJurusan.kdPendidikan, dbo.KualifikasiJurusan.StatusEnabled " & _
                     "FROM dbo.KualifikasiJurusan LEFT OUTER JOIN " & _
                     "dbo.Pendidikan ON dbo.KualifikasiJurusan.KdPendidikan = dbo.Pendidikan.KdPendidikan " & _
                     "WHERE (dbo.KualifikasiJurusan.StatusEnabled <> 0) OR (dbo.KualifikasiJurusan.StatusEnabled IS NULL)"
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgKualifikasi.DataSource = rs
            With dgKualifikasi
                .Columns(6).Width = 0
            End With
    End Select
End Sub

Private Function sp_JenisPendidikan(f_status As String) As Boolean
sp_JenisPendidikan = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisPendidikan", adVarChar, adParamInput, 3, txtKdJenisPendidikan.Text)
        .Parameters.Append .CreateParameter("JenisPendidikan", adVarChar, adParamInput, 30, Trim(txtJenisPendidikan.Text))
        .Parameters.Append .CreateParameter("OutputKdJenisPendidikan", adVarChar, adParamOutput, 3, Null)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisPendidikan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_JenisPendidikan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_Pendidikan(f_status As String) As Boolean
sp_Pendidikan = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdPendidikan", adVarChar, adParamInput, 4, txtKdPendidikan.Text)
        .Parameters.Append .CreateParameter("Pendidikan", adVarChar, adParamInput, 25, Trim(txtPendidikan.Text))
        .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 2, IIf(txtNoUrut.Text = "", Null, Trim(txtNoUrut.Text)))
        .Parameters.Append .CreateParameter("KdJenisPendidikan", adVarChar, adParamInput, 3, IIf(dcJenisPendidikan.Text = "", Null, Trim(dcJenisPendidikan.BoundText)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExt.Text = "", Null, Trim(txtKdExt.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 15, IIf(txtNamaExt.Text = "", Null, Trim(txtNamaExt.Text)))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_Pendidikan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_Pendidikan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Function sp_KualifikasiJurusan(f_status As String) As Boolean
sp_KualifikasiJurusan = True
Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdKualifikasiJurusan", adVarChar, adParamInput, 4, txtKdKualifikasi.Text)
        .Parameters.Append .CreateParameter("KualifikasiJurusan", adVarChar, adParamInput, 75, Trim(txtKualifikasi.Text))
        .Parameters.Append .CreateParameter("ReportDisplay", adVarChar, adParamInput, 75, IIf(txtRepDisplayKualifikasi.Text = "", Null, Trim(txtRepDisplayKualifikasi.Text)))
        .Parameters.Append .CreateParameter("KdPendidikan", adChar, adParamInput, 4, IIf(dcPendidikan.Text = "", Null, Trim(dcPendidikan.BoundText)))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, IIf(txtKdExtKualifikasi.Text = "", Null, Trim(txtKdExtKualifikasi.Text)))
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 75, IIf(txtNamaExtKualifikasi.Text = "", Null, Trim(txtNamaExtKualifikasi.Text)))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_status)
        
        .ActiveConnection = dbConn
        .CommandText = "AUD_KualifikasiJurusan"
        .CommandType = adCmdStoredProc
        .Execute
        
        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_KualifikasiJurusan = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
End Function

Private Sub cmdCancel_Click()
    Call clear
    Call SubLoadDCSource
    Call subLoadGridSource
    Call SSTab1_KeyPress(13)
End Sub

Private Sub cmdDel_Click()
On Error GoTo hell
        
    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtJenisPendidikan, "Jenis pendidikan kosong") = False Then Exit Sub
            If sp_JenisPendidikan("D") = False Then Exit Sub
        Case 1
            If Periksa("text", txtPendidikan, "Nama pendidikan kosong ") = False Then Exit Sub
            If sp_Pendidikan("D") = False Then Exit Sub
        Case 2
            If Periksa("text", txtKualifikasi, "Kualifikasi Jurusan kosong ") = False Then Exit Sub
            If sp_KualifikasiJurusan("D") = False Then Exit Sub
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
            If Periksa("text", txtJenisPendidikan, "Silahkan isi Jenis Pendidikan ") = False Then Exit Sub
            If sp_JenisPendidikan("A") = False Then Exit Sub

        Case 1
            If Periksa("text", txtPendidikan, "Silahkan isi Nama Pendidikan ") = False Then Exit Sub
            If sp_Pendidikan("A") = False Then Exit Sub

        Case 2
            If Periksa("text", txtKualifikasi, "Silahkan isi Kualifikasi Jurusan ") = False Then Exit Sub
            If sp_KualifikasiJurusan("A") = False Then Exit Sub
    End Select
    
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click
    
Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub dcJenisPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRepDisplay.SetFocus
End Sub

Private Sub dcPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKualifikasi.SetFocus
End Sub

Private Sub dgJenisPendidikan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If dgJenisPendidikan.ApproxCount = 0 Then Exit Sub
    txtKdJenisPendidikan.Text = dgJenisPendidikan.Columns(0).Value
    txtJenisPendidikan.Text = dgJenisPendidikan.Columns(1).Value
End Sub

Private Sub dgKualifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKualifikasi.SetFocus
End Sub

Private Sub dgKualifikasi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errLoad
    If dgKualifikasi.ApproxCount = 0 Then Exit Sub
    dcPendidikan.BoundText = dgKualifikasi.Columns(6)
    txtKdKualifikasi = dgKualifikasi.Columns(0)
    txtKualifikasi = dgKualifikasi.Columns(1)
    txtRepDisplayKualifikasi = dgKualifikasi.Columns(3)
    txtKdExtKualifikasi.Text = dgKualifikasi.Columns(4)
    txtNamaExtKualifikasi.Text = dgKualifikasi.Columns(5)
    
Exit Sub
errLoad:
End Sub

Private Sub dgPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPendidikan.SetFocus
End Sub

Private Sub dgPendidikan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error GoTo errLoad
    If dgPendidikan.ApproxCount = 0 Then Exit Sub
    txtKdPendidikan.Text = dgPendidikan.Columns(0)
    txtPendidikan.Text = dgPendidikan.Columns(1)
    If IsNull(dgPendidikan.Columns(2)) Then txtNoUrut.Text = "" Else txtNoUrut.Text = dgPendidikan.Columns(2)
    If IsNull(dgPendidikan.Columns(6)) Then dcJenisPendidikan.BoundText = "" Else dcJenisPendidikan.BoundText = dgPendidikan.Columns(6)
    If IsNull(dgPendidikan.Columns(4)) Then txtKdExt.Text = "" Else txtKdExt.Text = dgPendidikan.Columns(4)
    If IsNull(dgPendidikan.Columns(5)) Then txtNamaExt.Text = "" Else txtNamaExt.Text = dgPendidikan.Columns(5)
Exit Sub
errLoad:
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

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call SubLoadDCSource
    Call subLoadGridSource
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
On Error GoTo errLoad
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtPendidikan.SetFocus
            Case 1
                dcPendidikan.SetFocus
        End Select
    End If
errLoad:
End Sub

Private Sub txtJenisPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExt.SetFocus
End Sub

Private Sub txtKdExtKualifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNamaExtKualifikasi.SetFocus
End Sub

Private Sub txtKdPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPendidikan.SetFocus
End Sub

Private Sub txtKdKualifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKualifikasi.SetFocus
End Sub

Private Sub txtKualifikasi_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            dgKualifikasi.SetFocus
    End Select
End Sub

Private Sub txtKualifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRepDisplayKualifikasi.SetFocus
End Sub

Private Sub txtNoUrut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisPendidikan.SetFocus
End Sub

Private Sub txtPendidikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNoUrut.SetFocus
End Sub

Private Sub txtNamaExtKualifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtRepDisplay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtRepDisplayKualifikasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtKualifikasi.SetFocus
End Sub
