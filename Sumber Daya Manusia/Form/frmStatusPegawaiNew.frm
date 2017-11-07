VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmStatusPegawaiNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Master Status Pegawai"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatusPegawaiNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   7110
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   360
      TabIndex        =   43
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   7800
      Width           =   1215
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6450
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   6900
      _ExtentX        =   12171
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
      TabCaption(0)   =   "Status"
      TabPicture(0)   =   "frmStatusPegawaiNew.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Jenis Status"
      TabPicture(1)   =   "frmStatusPegawaiNew.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Alasan Status"
      TabPicture(2)   =   "frmStatusPegawaiNew.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
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
         Height          =   5655
         Left            =   -74760
         TabIndex        =   20
         Top             =   600
         Width           =   6360
         Begin VB.TextBox txtKdExt2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   40
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chkSts2 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   4800
            TabIndex        =   39
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNmExt2 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   38
            Top             =   1680
            Width           =   4575
         End
         Begin VB.TextBox txtAlasan 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   22
            Top             =   600
            Width           =   4575
         End
         Begin VB.TextBox txtKdAlasan 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   21
            Top             =   240
            Width           =   720
         End
         Begin MSDataGridLib.DataGrid dgAlasan 
            Height          =   3330
            Left            =   255
            TabIndex        =   23
            Top             =   2160
            Width           =   5880
            _ExtentX        =   10372
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
         Begin MSDataListLib.DataCombo dcStatusAlasan 
            Height          =   315
            Left            =   1560
            TabIndex        =   24
            Top             =   960
            Width           =   2880
            _ExtentX        =   5080
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   210
            Index           =   3
            Left            =   240
            TabIndex        =   42
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   41
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   210
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Alasan Status"
            Height          =   210
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            Height          =   210
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   420
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
         TabIndex        =   12
         Top             =   600
         Width           =   6360
         Begin VB.TextBox txtKdExt1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   35
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CheckBox chkSts1 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   4800
            TabIndex        =   34
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNmExt1 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1680
            Width           =   4575
         End
         Begin VB.TextBox txtKdJenisStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtJenisStatus 
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
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   13
            Top             =   600
            Width           =   4575
         End
         Begin MSDataGridLib.DataGrid dgJenisStatus 
            Height          =   3330
            Left            =   240
            TabIndex        =   15
            Top             =   2160
            Width           =   5880
            _ExtentX        =   10372
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
         Begin MSDataListLib.DataCombo dcStatusJenis 
            Height          =   315
            Left            =   1560
            TabIndex        =   16
            Top             =   960
            Width           =   2880
            _ExtentX        =   5080
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   1320
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   1680
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Jenis Status"
            Height          =   210
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Kode"
            Height          =   210
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Status"
            Height          =   210
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   525
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5655
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   6375
         Begin VB.TextBox txtKdExt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   30
            Top             =   960
            Width           =   1815
         End
         Begin VB.CheckBox chkSts 
            Alignment       =   1  'Right Justify
            Caption         =   "Status Aktif"
            Height          =   255
            Left            =   4800
            TabIndex        =   29
            Top             =   960
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox txtNmExt 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   28
            Top             =   1320
            Width           =   4575
         End
         Begin VB.TextBox txtKdStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtStatus 
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
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   7
            Top             =   600
            Width           =   4575
         End
         Begin MSDataGridLib.DataGrid dgStatus 
            Height          =   3690
            Left            =   255
            TabIndex        =   9
            Top             =   1800
            Width           =   5880
            _ExtentX        =   10372
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
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Kode External"
            Height          =   210
            Index           =   6
            Left            =   240
            TabIndex        =   32
            Top             =   960
            Width           =   1140
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Nama External"
            Height          =   210
            Index           =   7
            Left            =   240
            TabIndex        =   31
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nama Status"
            Height          =   210
            Left            =   240
            TabIndex        =   11
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Kode "
            Height          =   210
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   480
         End
      End
   End
   Begin VB.Image Image4 
      Height          =   945
      Left            =   5280
      Picture         =   "frmStatusPegawaiNew.frx":0D1E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmStatusPegawaiNew.frx":1AA6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9495
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmStatusPegawaiNew.frx":3104
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   7080
      Picture         =   "frmStatusPegawaiNew.frx":5AC5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmStatusPegawaiNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sStatus As String

Sub clear()
    On Error Resume Next

    Select Case SSTab1.Tab
        Case 0
            txtKdStatus.Text = ""
            txtstatus.Text = ""
            txtKdExt.Text = ""
            chkSts.Value = 1
            txtNmExt.Text = ""
        Case 1
            txtKdJenisStatus.Text = ""
            txtJenisStatus.Text = ""
            dcStatusJenis.BoundText = ""
            sStatus = "A"
            txtKdExt1.Text = ""
            chkSts1.Value = 1
            txtNmExt1.Text = ""
        Case 2
            dcStatusAlasan.BoundText = ""
            txtKdAlasan.Text = ""
            txtalasan.Text = ""
            sStatus = "A"
            txtKdExt2.Text = ""
            chkSts2.Value = 1
            txtNmExt2.Text = ""
    End Select
End Sub

Sub subLoadDcSource()
    Select Case SSTab1.Tab
        Case 1
            Call msubDcSource(dcStatusJenis, rs, "SELECT KdStatus, Status FROM StatusPegawai where StatusEnabled='1' order by Status")
        Case 2
            Call msubDcSource(dcStatusAlasan, rs, "SELECT KdStatus, Status FROM StatusPegawai where StatusEnabled='1' order by Status")
    End Select
End Sub

Sub subLoadGridSource()
    Select Case SSTab1.Tab
        Case 0
            strSQL = "SELECT KdStatus as Kode, Status,KodeExternal,NamaExternal,StatusEnabled FROM StatusPegawai order by Kode"
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgStatus.DataSource = rs
            With dgStatus
                .Columns(0).Width = 1000
                .Columns(1).Width = 4000
                .Columns(4).Width = 1250
            End With

        Case 1
            strSQL = "SELECT dbo.JenisStatusPegawai.KdJenisStatus as Kode, dbo.JenisStatusPegawai.JenisStatus as [Jenis Status], dbo.JenisStatusPegawai.KdStatus, dbo.StatusPegawai.Status," & _
            " dbo.JenisStatusPegawai.KodeExternal,dbo.JenisStatusPegawai.NamaExternal,dbo.JenisStatusPegawai.StatusEnabled FROM dbo.JenisStatusPegawai INNER JOIN" & _
            " dbo.StatusPegawai ON dbo.JenisStatusPegawai.KdStatus = dbo.StatusPegawai.KdStatus "
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgJenisStatus.DataSource = rs
            With dgJenisStatus
                .Columns(0).Width = 500
                .Columns(1).Width = 2800
                .Columns(2).Width = 0
                .Columns(3).Width = 2200
                .Columns(6).Width = 1250
            End With

        Case 2
            strSQL = "SELECT dbo.AlasanStatusPegawai.KdAlasanStatus as Kode, dbo.AlasanStatusPegawai.AlasanStatus as [Alasan Status], dbo.AlasanStatusPegawai.KdStatus, " & _
            "dbo.StatusPegawai.Status, dbo.AlasanStatusPegawai.KodeExternal,dbo.AlasanStatusPegawai.NamaExternal,dbo.AlasanStatusPegawai.StatusEnabled" & _
            " FROM dbo.AlasanStatusPegawai INNER JOIN " & _
            "dbo.StatusPegawai ON dbo.AlasanStatusPegawai.KdStatus = dbo.StatusPegawai.KdStatus"
            Set rs = Nothing
            Call msubRecFO(rs, strSQL)
            Set dgAlasan.DataSource = rs
            With dgAlasan
                .Columns(0).Width = 1000
                .Columns(1).Width = 2200
                .Columns(2).Width = 0
                .Columns(3).Width = 2200
                .Columns(6).Width = 1250
            End With
    End Select
End Sub

Private Function sp_Status(f_Status As String) As Boolean
On Error GoTo ErrSpStatus
    sp_Status = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, txtKdStatus.Text)
        .Parameters.Append .CreateParameter("Status", adVarChar, adParamInput, 30, Trim(txtstatus.Text))
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExt.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExt.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
        .Parameters.Append .CreateParameter("StatusCode", adChar, adParamInput, 1, f_Status)

        .ActiveConnection = dbConn
        .CommandText = "AUD_StatusPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_Status = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
  Exit Function
ErrSpStatus:
    If f_Status = "D" Then
            MsgBox "Data tidak bisa di hapus, data sudah di pakai", vbCritical
    Else
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
    End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_Status = False
End Function

Private Function sp_JenisStatus(f_Status As String) As Boolean
On Error GoTo ErrspJenisStatus
    sp_JenisStatus = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdJenisStatus", adTinyInt, adParamInput, , IIf(txtKdJenisStatus.Text = "", Null, txtKdJenisStatus))
        .Parameters.Append .CreateParameter("JenisStatus", adVarChar, adParamInput, 50, Trim(txtJenisStatus.Text))
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatusJenis.BoundText)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExt1.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExt1.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts1.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("output", adTinyInt, adParamInput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_JenisStatusPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_JenisStatus = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
ErrspJenisStatus:
   If f_Status = "D" Then
            MsgBox "Data tidak bisa di hapus, data sudah di pakai", vbCritical
    Else
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
    End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_JenisStatus = False

End Function

Private Function sp_AlasanStatus(f_Status As String) As Boolean
On Error GoTo ErrspAlasanStatus
    sp_AlasanStatus = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdAlasanStatus", adTinyInt, adParamInput, , IIf(txtKdAlasan.Text = "", Null, txtKdAlasan.Text))
        .Parameters.Append .CreateParameter("AlasanStatus", adVarChar, adParamInput, 50, Trim(txtalasan.Text))
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatusAlasan.BoundText)
        .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, txtKdExt2.Text)
        .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, txtNmExt2.Text)
        .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts2.Value)
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)
        .Parameters.Append .CreateParameter("output", adTinyInt, adParamInput, , Null)

        .ActiveConnection = dbConn
        .CommandText = "AUD_AlasanStatusPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            sp_AlasanStatus = False
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
Exit Function
ErrspAlasanStatus:
   If f_Status = "D" Then
            MsgBox "Data tidak bisa di hapus, data sudah di pakai", vbCritical
    Else
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
    End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
        sp_AlasanStatus = False
    
End Function

Private Sub cmdCancel_Click()
    Call clear
    Call subLoadDcSource
    Call subLoadGridSource
    Call SSTab1_KeyPress(13)
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Select Case SSTab1.Tab
        Case 0
            If dgStatus.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakStatus.Show
        Case 1
            If dgJenisStatus.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakJenisStatus.Show
        Case 2
            If dgAlasan.ApproxCount = 0 Then Exit Sub
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmCetakAlasanStatus.Show
    End Select
hell:
End Sub

Private Sub cmdDel_Click()
    On Error GoTo hell

    If MsgBox("Yakin akan menghapus data ini ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtstatus, "Nama status kosong") = False Then Exit Sub
            If sp_Status("D") = False Then Exit Sub
        Case 1
            If Periksa("text", txtJenisStatus, "Nama jenis status ") = False Then Exit Sub
            If Periksa("datacombo", dcStatusJenis, "jenis status kosong ") = False Then Exit Sub
            If sp_JenisStatus("D") = False Then Exit Sub
        Case 2
            If Periksa("text", txtalasan, "Nama alasan kosong ") = False Then Exit Sub
            If Periksa("datacombo", dcStatusAlasan, "alasan status kosong ") = False Then Exit Sub
            If sp_AlasanStatus("D") = False Then Exit Sub
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
    On Error GoTo Errload

    Select Case SSTab1.Tab
        Case 0
            If Periksa("text", txtstatus, "Silahkan isi nama status ") = False Then Exit Sub
            If sp_Status("A") = False Then Exit Sub

        Case 1
            If Periksa("text", txtJenisStatus, "Silahkan isi nama jenis status ") = False Then Exit Sub
            If Periksa("datacombo", dcStatusJenis, "Silahkan isi nama status ") = False Then Exit Sub
            If sp_JenisStatus(sStatus) = False Then Exit Sub

        Case 2
            If Periksa("text", txtalasan, "Silahkan isi alasan status ") = False Then Exit Sub
            If Periksa("datacombo", dcStatusAlasan, "Silahkan isi nama status ") = False Then Exit Sub
            If sp_AlasanStatus(sStatus) = False Then Exit Sub
    End Select

    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call cmdCancel_Click

    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcStatusJenis_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then txtKdExt1.SetFocus
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcStatusJenis.Text)) = 0 Then txtKdExt1.SetFocus: Exit Sub
        If dcStatusJenis.MatchedWithList = True Then txtKdExt1.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdStatus, Status FROM StatusPegawai WHERE Status LIKE '%" & dcStatusJenis.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcStatusJenis.BoundText = dbRst(0).Value
        dcStatusJenis.Text = dbRst(1).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dcStatusAlasan_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then txtKdExt2.SetFocus
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcStatusAlasan.Text)) = 0 Then txtKdExt2.SetFocus: Exit Sub
        If dcStatusAlasan.MatchedWithList = True Then txtKdExt2.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdStatus, Status FROM StatusPegawai WHERE Status LIKE '%" & dcStatusAlasan.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcStatusAlasan.BoundText = dbRst(0).Value
        dcStatusAlasan.Text = dbRst(1).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgAlasan_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgAlasan
    WheelHook.WheelHook dgAlasan
End Sub

Private Sub dgJenisStatus_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgJenisStatus
    WheelHook.WheelHook dgJenisStatus
End Sub

Private Sub dgStatus_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgStatus
    WheelHook.WheelHook dgStatus
End Sub

Private Sub dgStatus_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatus.ApproxCount = 0 Then Exit Sub
    txtKdStatus.Text = dgStatus.Columns(0).Value
    txtstatus.Text = dgStatus.Columns(1).Value
    txtKdExt.Text = dgStatus.Columns(2).Value
    txtNmExt.Text = dgStatus.Columns(3).Value
    chkSts.Value = dgStatus.Columns(4).Value
    sStatus = "U"
End Sub

Private Sub dgAlasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtalasan.SetFocus
End Sub

Private Sub dgAlasan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Errload
    If dgAlasan.ApproxCount = 0 Then Exit Sub
    dcStatusAlasan.BoundText = dgAlasan.Columns(2)
    txtKdAlasan = dgAlasan.Columns(0)
    txtalasan = dgAlasan.Columns(1)
    txtKdExt2.Text = dgAlasan.Columns(4).Value
    txtNmExt2.Text = dgAlasan.Columns(5).Value
    chkSts2.Value = dgAlasan.Columns(6).Value
    sStatus = "U"
    Exit Sub
Errload:
End Sub

Private Sub dgJenisStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisStatus.SetFocus
End Sub

Private Sub dgJenisStatus_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo Errload
    If dgJenisStatus.ApproxCount = 0 Then Exit Sub
    txtKdJenisStatus.Text = dgJenisStatus.Columns(0)
    txtJenisStatus.Text = dgJenisStatus.Columns(1)
    dcStatusJenis.BoundText = dgJenisStatus.Columns(2)
    txtKdExt1.Text = dgJenisStatus.Columns(4).Value
    txtNmExt1.Text = dgJenisStatus.Columns(5).Value
    chkSts1.Value = dgJenisStatus.Columns(6).Value
    sStatus = "U"
    Exit Sub
Errload:
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
    Call cmdCancel_Click
End Sub

Private Sub SSTab1_KeyPress(KeyAscii As Integer)
    On Error GoTo Errload
    If KeyAscii = 13 Then
        Select Case SSTab1.Tab
            Case 0
                txtstatus.SetFocus
            Case 1
                txtJenisStatus.SetFocus
            Case 2
                txtalasan.SetFocus
        End Select
    End If
Errload:
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt.SetFocus
End Sub

Private Sub txtKdExt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt1.SetFocus
End Sub

Private Sub txtKdExt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt2.SetFocus
End Sub

Private Sub txtNmExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNmExt1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtNmExt2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSave.SetFocus
End Sub

Private Sub txtstatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtKdJenisStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtJenisStatus.SetFocus
End Sub

Private Sub txtKdAlasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtalasan.SetFocus
End Sub

Private Sub txtAlasan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusAlasan.SetFocus
End Sub

Private Sub txtJenisStatus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcStatusJenis.SetFocus
End Sub

