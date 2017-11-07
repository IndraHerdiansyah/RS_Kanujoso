VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetailRequest 
   Caption         =   "Kategori"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmDetailRequest.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   10260
   Begin VB.CommandButton cmdEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtKdTask 
      Height          =   285
      Left            =   6360
      TabIndex        =   53
      Top             =   120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame5"
      Height          =   495
      Left            =   6240
      TabIndex        =   46
      Top             =   9840
      Width           =   2415
      Begin VB.CommandButton btnNext 
         Height          =   315
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Next 250"
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton btnLast 
         Height          =   315
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Last 250"
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton btnPrev 
         Height          =   315
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Previous 250"
         Top             =   120
         Width           =   315
      End
      Begin VB.CommandButton btnFirst 
         Height          =   315
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "First 250"
         Top             =   120
         Width           =   315
      End
      Begin VB.Label lblPageInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0 - 0 of 0"
         Height          =   255
         Left            =   -1800
         TabIndex        =   52
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label lblCurrentRecord 
         AutoSize        =   -1  'True
         Height          =   315
         Left            =   360
         TabIndex        =   51
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   44
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton txtCariProblem 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Cari Problem"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12840
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdLoad 
      Height          =   375
      Left            =   5040
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame fratambah 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Tambah Riwayat "
      Height          =   5895
      Left            =   2400
      TabIndex        =   6
      Top             =   10800
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   4800
         TabIndex        =   71
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid fgPerawatPerPelayanan 
         Height          =   1335
         Left            =   1440
         TabIndex        =   18
         Top             =   3720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2355
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   8577768
         ForeColorFixed  =   -2147483627
         ForeColorSel    =   -2147483628
         BackColorBkg    =   16777215
         FocusRect       =   0
         HighLight       =   2
         FillStyle       =   1
         GridLines       =   3
         SelectionMode   =   1
         AllowUserResizing=   1
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
      End
      Begin VB.TextBox txtproses1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   19
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtPetugas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   17
         Top             =   480
         Width           =   3735
      End
      Begin VB.CommandButton cmdSimpanRiwayat 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdBatalRiwayat 
         Caption         =   "Batal"
         Height          =   375
         Left            =   8640
         TabIndex        =   14
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdTutup 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   10080
         TabIndex        =   13
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox TxtJobIdent 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   3840
         Width           =   2295
      End
      Begin VB.CheckBox chkPerawat 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Petugas SIM"
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   7560
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CommandButton Comman 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   4200
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSComctlLib.ListView lvPemeriksa 
         Height          =   1815
         Left            =   7560
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nama Pemeriksa"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtMulaiJob 
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
         Format          =   129302531
         UpDown          =   -1  'True
         CurrentDate     =   40331
      End
      Begin MSComCtl2.DTPicker dtSelesaiJob 
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
         Format          =   129302531
         UpDown          =   -1  'True
         CurrentDate     =   40331
      End
      Begin MSDataListLib.DataCombo dcStatus3 
         Height          =   360
         Left            =   7560
         TabIndex        =   62
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtSolusi 
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
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   11295
      End
      Begin VB.Label row 
         Caption         =   "0"
         Height          =   255
         Left            =   5880
         TabIndex        =   70
         Top             =   3840
         Width           =   255
      End
      Begin VB.Label lbbantu 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Pelaksana"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6360
         TabIndex        =   29
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "TglSelesai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Mulai"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Job Identity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Label19"
         Height          =   255
         Left            =   10200
         TabIndex        =   25
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Proses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6360
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   8760
         TabIndex        =   23
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   6360
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdTambahProblem 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Caption         =   "Tambah ke Problem"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   14895
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Height          =   2295
         Left            =   240
         TabIndex        =   54
         Top             =   120
         Width           =   4695
         Begin MSDataListLib.DataCombo dgKategori 
            Height          =   360
            Left            =   2160
            TabIndex        =   55
            Top             =   840
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcPrioritas 
            Height          =   360
            Left            =   2160
            TabIndex        =   59
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcStatus 
            Height          =   360
            Left            =   2160
            TabIndex        =   63
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Kategori"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Prioritas"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   1320
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000E&
         Height          =   2295
         Left            =   9960
         TabIndex        =   37
         Top             =   120
         Width           =   4695
         Begin MSComCtl2.DTPicker DtpAwal 
            Height          =   375
            Left            =   1680
            TabIndex        =   38
            Top             =   840
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   129236995
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker DtpAkhir 
            Height          =   375
            Left            =   1680
            TabIndex        =   39
            Top             =   1320
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   129236995
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin MSComCtl2.DTPicker dtpOrder 
            Height          =   375
            Left            =   1680
            TabIndex        =   64
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
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
            CustomFormat    =   "dd MMM yyyy HH:mm"
            Format          =   129236995
            UpDown          =   -1  'True
            CurrentDate     =   38209
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Penanggung Jawab"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblPenanggungJawab 
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   375
            Left            =   2040
            TabIndex        =   42
            Top             =   1800
            Width           =   3375
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Mulai"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Selesai"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         Height          =   2295
         Left            =   5040
         TabIndex        =   34
         Top             =   120
         Width           =   4695
         Begin MSDataListLib.DataCombo dcRuangan 
            Height          =   360
            Left            =   1320
            TabIndex        =   60
            Top             =   720
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcPetugas 
            Height          =   360
            Left            =   1320
            TabIndex        =   61
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcTingkatKesulitan 
            Height          =   360
            Left            =   1320
            TabIndex        =   66
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tingkat Kesulitan"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Ruangan"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   975
         End
      End
   End
   Begin VB.TextBox txtmasalah 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   14775
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "Permintaan Maintenance/Perbaikan"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   4320
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000C000&
      Caption         =   "Permintaan Perubahan System/Aplikasi"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton cmdTambah 
      BackColor       =   &H0000C000&
      Caption         =   "Tambah Riwayat"
      Enabled         =   0   'False
      Height          =   375
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvRiwayatJob 
      Height          =   3435
      Left            =   240
      TabIndex        =   45
      Top             =   4680
      Width           =   14940
      _ExtentX        =   26353
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No Permintaan"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "No. Pelaksanaan"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Resolusi"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Pelaksana"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tanggal Mulai"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tanggal Selesai"
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.Label lblNoRequest 
      BackStyle       =   0  'Transparent
      Caption         =   "Request ID"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Request ID"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   9720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   15000
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmDetailRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdmulai As Date
Dim mdselesai As Date
Dim subKdPemeriksa() As String
Dim subJmlTotal As Integer
Dim j As Integer
Dim kedit As String
Dim x As ListItem

Private Sub cmdEdit_Click()
    With FrmTaskList
        FrmTaskList.txtKdTask = txtKdTask.Text
        If dcStatus.Text = "Dalam Proses" Then
            dcStatus.BoundText = "04"
            dcStatus.Enabled = False
        End If
        strSQL = "select * from V_SimOrder where KdTask = '" & txtKdTask.Text & "'"
        Set rs = Nothing
        Call msubRecFO(rs, strSQL)

        .txtmasalah.Text = rs.Fields("Masalah").Value
        .dtOrder.Value = rs.Fields("tglOrder").Value
        .dcRuangan.BoundText = rs.Fields("KdRuangan").Value
        lblPenanggungJawab.Caption = rs.Fields("IdPelapor").Value
        .txtmasalah.Enabled = False
        .dtOrder.Enabled = False
        .dcRuangan.Enabled = False
    End With
    FrmTaskList.Show

End Sub

Private Sub Command4_Click()
    Call subLoadPelayananPerPerawat
End Sub

Private Sub dtMulaiJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtSelesaiJob.SetFocus
    End If
End Sub

Private Sub dtSelesaiJob_KeyDown(KeyCode As Integer, Shift As Integer)
    If kecode = vbKeyReturn Then
        txtPetugas.SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ListJob()

    strSQL = "select distinct KdTask, KdJob, Solusi, TglMulai, TglSelesai from SIMTask where KdTask = '" & strNoRequest & "' order by TglMulai"
    Set rs = Nothing
    rs.Open strSQL, dbConn, adOpenStatic, adLockReadOnly
    If rs.EOF = True Or rs.BOF = True Then
        Exit Sub
    End If
    rs.MoveFirst
    With lvRiwayatJob
        For i = 1 To rs.RecordCount
            Set x = lvRiwayatJob.ListItems.add(, , rs.Fields(0).Value)
            x.SubItems(1) = rs.Fields(1).Value
            x.SubItems(2) = rs.Fields(2).Value
            x.SubItems(3) = "-"
            x.SubItems(4) = rs.Fields(3).Value
            x.SubItems(5) = rs.Fields(4).Value
            rs.MoveNext
        Next i
    End With
End Sub

Private Sub subDcSource()
    Call msubDcSource(dcRuangan, rs, "select KdRuangan, NamaRuangan from Ruangan")
    Call msubDcSource(dcPetugas, rs, "Select IdPegawai, NamaLengkap from DataPegawai")
    Call msubDcSource(dcPrioritas, rs, "Select KdPrioritas, NamaPrioritas from SIMPrioritas")
    Call msubDcSource(dgKategori, rs, "Select KdKategori, NamaKategori from SIMKategori")
    Call msubDcSource(dcTingkatKesulitan, rs, "Select KdTingkatKesulitan, NamaTingkatKesulitan from SIMTingkatKesulitan")

    Call msubDcSource(dcStatus, rs, "Select KdStatus, NamaStatus from SIMStatus")
    Call msubDcSource(dcStatus3, rs, "Select KdStatus, NamaStatus from SIMStatus")
End Sub

Private Sub cmdLoad_Click()
    strSQL = "Select KdTask, KdRuangan, Masalah, Pengirim, NamaRuangan, IdPelapor, TglOrder, TglMulai, TglSelesai, KdKategori, Kategori, Proses, KdPrioritas, Prioritas, KdStatus, Status,  KdInstalasi, IdPegawai, KdRuanganPelaksana, PenanggungJawab from V_simOrder where KdTask = '" & strNoRequest & "'"
    Set rs = Nothing
    Call msubRecFO(rs, strSQL)

    txtmasalah.Text = rs.Fields("Masalah")
    If IsNull(rs.Fields("Status")) = True Then
        dcStatus.Text = ""
    Else
        If rs.Fields("KdStatus").Value = "01" Then
            dcStatus.BoundText = "06"
        ElseIf rs.Fields("Status").Value = "06" Then
            dcStatus.BoundText = "02"
        Else
            dcStatus.BoundText = rs.Fields("KdStatus").Value
        End If
    End If

    If IsNull(rs.Fields("Kategori")) = True Then
        dgKategori.Text = ""
    Else
        dgKategori.Text = rs.Fields("Kategori")
    End If

    If IsNull(rs.Fields("Prioritas")) = True Then
        dcPrioritas.Text = ""
    Else
        dcPrioritas.Text = rs.Fields("Prioritas")
    End If

    If IsNull(rs.Fields("Pengirim")) = True Then
        dcPetugas.Text = ""
    Else
        dcPetugas.Text = rs.Fields("Pengirim")
    End If

    If IsNull(rs.Fields("NamaRuangan")) = True Then
        dcRuangan.Text = ""
    Else
        dcRuangan.Text = rs.Fields("NamaRuangan")
    End If

    If IsNull(rs.Fields("TglOrder")) = True Then
        dtpOrder.Value = ""
    Else
        dtpOrder.Value = rs.Fields("TglOrder")
    End If

    If IsNull(rs.Fields("tglmulai").Value) = True Then
        DtpAwal.Value = Now
    Else
        DtpAwal.Value = rs.Fields("tglMulai")
    End If
    If IsNull(rs.Fields("TglSelesai").Value) = True Then
        DtpAkhir.Value = Now
    Else
        DtpAkhir.Value = rs.Fields("TglSelesai")
    End If

    If IsNull(rs.Fields("PenanggungJawab").Value) = True Then
        lblPenanggungJawab.Caption = ""
    Else
        lblPenanggungJawab.Caption = rs.Fields("PenanggungJawab")
    End If
    txtKdTask.Text = strNoRequest

    dcStatus.Enabled = False
    dgKategori.Enabled = False
    dcPrioritas.Enabled = False
    dcPetugas.Enabled = False
    dcRuangan.Enabled = False
    dcTingkatKesulitan.Enabled = False

    ListJob
End Sub

Private Sub cmdSimpanRiwayat_Click()
    Call Command4_Click
    If sp_Job() = False Then Exit Sub
    If kedit = "1" Then Exit Sub

    Set rs = Nothing
    strSQL = "select KdJob from SimTask where KdTask = '" & txtKdTask.Text & "' and KdJob = '" & TxtJobIdent.Text & "'"
    j = fgPerawatPerPelayanan.Rows - 1

    Dim N As Integer
    For i = 1 To fgPerawatPerPelayanan.Rows - 1
        With fgPerawatPerPelayanan
            fgPerawatPerPelayanan.row = 1

            If sp_PetugasperJOD() = False Then Exit Sub
        End With

        If fgPerawatPerPelayanan.row <> j Then
            fgPerawatPerPelayanan.row = fgPerawatPerPelayanan.row + 1
        End If

    Next i

    dtMulaiJob.Value = Now
    dtSelesaiJob.Value = Now
    txtPetugas.Text = ""
    txtSolusi.Text = ""
    dcStatus3.Text = ""
    TxtJobIdent.Text = ""
    fgPerawatPerPelayanan.clear
    fgPerawatPerPelayanan.Refresh
    lvPemeriksa.Refresh

    Call MsgBox("Data Pelaksana Berhasil di Simpan", vbOKOnly, "SIM-JOD Validation")
    cmdTutup.SetFocus
    dtMulaiJob.Value = Now
    dtSelesaiJob.Value = Now
    fgPerawatPerPelayanan.clear
    chkPerawat.Caption = "Petugas SIM"
    txtproses1.Text = ""
    ListJob

End Sub

Private Sub cmdTambahRiwayat_Click()
    frmInputJOD.txtKdTask = txtKdTask.Text

    frmInputJOD.Show
    frmInputJOD.txtKdTask.Enabled = False

End Sub

Private Function sp_Task() As Boolean
    On Error GoTo errLoad
    sp_StrukTerima = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        If Len(Trim(txtKdTask.Text)) = 0 Then
            .Parameters.Append .CreateParameter("KdTask", adChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("KdTask", adChar, adParamInput, 10, txtKdTask.Text)
        End If
        .Parameters.Append .CreateParameter("KdRuanganAsal", adChar, adParamInput, 3, dcRuang.BoundText)
        .Parameters.Append .CreateParameter("IdPelapor", adChar, adParamInput, 10, dcPetugas.BoundText)
        .Parameters.Append .CreateParameter("Masalah", adChar, adParamInput, 3000, txtmasalah.Text)
        .Parameters.Append .CreateParameter("TglOrder", adDate, adParamInput, , Format(dtOrder.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("KdKategory", adChar, adParamInput, 2, dgKategori.BoundText)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatus.BoundText)
        .Parameters.Append .CreateParameter("KdPrioritas", adChar, adParamInput, 2, dcPrioritas.BoundText)
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, strIDPegawaiAktif)
        .Parameters.Append .CreateParameter("KdRuanganPelaksana", adChar, adParamInput, 3, mstrKdRuangan)
        .Parameters.Append .CreateParameter("OutputKdTask", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "SIMAdd_Task"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Task = False
        Else

            txtKdTask.Text = .Parameters("OutputKdTask").Value
            Call MsgBox("Data Sudah Tersimpan", vbOKOnly, "PERHATIAN")

        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_Task = False
    Call deleteADOCommandParameters(dbcmd)

    Call msubPesanError
End Function

Private Function sp_Job() As Boolean
    On Error GoTo errLoad
    sp_Job = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        If Len(Trim(TxtJobIdent.Text)) = 0 Then
            .Parameters.Append .CreateParameter("KdJob", adVarChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("KdJob", adVarChar, adParamInput, 10, TxtJobIdent.Text)
        End If
        .Parameters.Append .CreateParameter("KdTask", adVarChar, adParamInput, 10, txtKdTask.Text)
        .Parameters.Append .CreateParameter("tglMulai", adDate, adParamInput, , Format(dtMulaiJob.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("tglSelesai", adDate, adParamInput, , Format(dtSelesaiJob.Value, "yyyy/MM/dd HH:mm:ss"))
        .Parameters.Append .CreateParameter("Solusi", adVarChar, adParamInput, 3000, txtSolusi.Text)
        .Parameters.Append .CreateParameter("KdStatus", adChar, adParamInput, 2, dcStatus3.BoundText)
        .Parameters.Append .CreateParameter("Proses", adInteger, adParamInput, , Val(txtproses1.Text))
        .Parameters.Append .CreateParameter("OutputKdJobTemp", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "SIMAU_Job"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_Job = False
        Else
            TxtJobIdent.Text = .Parameters("OutputKdJobTemp").Value
            Call MsgBox("Data Sudah Tersimpan", vbOKOnly, "PERHATIAN")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_Job = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError

End Function

Private Function sp_PetugasperJOD() As Boolean
    On Error GoTo errLoad
    sp_PetugasperJOD = True
    Set dbcmd = New ADODB.Command
    With dbcmd
        .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
        .Parameters.Append .CreateParameter("KdTask", adVarChar, adParamInput, 10, txtKdTask.Text)

        If Len(Trim(TxtJobIdent.Text)) = 0 Then
            .Parameters.Append .CreateParameter("KdJob", adVarChar, adParamInput, 10, Null)
        Else
            .Parameters.Append .CreateParameter("KdJob", adVarChar, adParamInput, 10, TxtJobIdent.Text)
        End If

        .Parameters.Append .CreateParameter("IdPetugas", adChar, adParamInput, 10, fgPerawatPerPelayanan.TextMatrix(fgPerawatPerPelayanan.row, 1))
        .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, "A")
        .Parameters.Append .CreateParameter("OutputIdPegawai", adChar, adParamOutput, 10, Null)

        .ActiveConnection = dbConn
        .CommandText = "SIM_AUPetugas"
        .CommandType = adCmdStoredProc
        .Execute

        If .Parameters("return_value").Value <> 0 Then
            MsgBox "Error - Ada kesalahan dalam penyimpanan data struk terima, Hubungi administrator", vbCritical, "Error"
            sp_PetugasperJOD = False
        Else
            Call MsgBox("Data Sudah Tersimpan", vbOKOnly, "PERHATIAN")
        End If
        Call deleteADOCommandParameters(dbcmd)
        Set dbcmd = Nothing
    End With
    Exit Function
errLoad:
    sp_PetugasperJOD = False
    Call deleteADOCommandParameters(dbcmd)
    Set dbcmd = Nothing
    Call msubPesanError

End Function

Private Sub cmdTambah_Click()
    dtMulaiJob.Value = Now
    dtSelesaiJob.Value = Now

    fratambah.Visible = True
End Sub

Private Sub cmdTutupRiwayat_Click()
    fratambah.Visible = False
End Sub

Private Sub txtnorequest_Change()

End Sub

Private Sub cmdTutup_Click()
    fratambah.Visible = False
End Sub

Private Sub Command5_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Me.WindowState = 2
    subDcSource
    ListJob
    Call cmdLoad_Click
End Sub

Private Sub Frame6_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub txtKdTask_Change()
    lblNoRequest.Caption = txtKdTask.Text
End Sub

Private Sub txtPetugas_Change()
    On Error GoTo errLoad

    Call subLoadListPemeriksa("where v_DaftarPemeriksaPasien.[Nama Pemeriksa] LIKE '%" & txtPetugas.Text & "%'")
    lvPemeriksa.Visible = True

    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Private Sub lvPemeriksa_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim blnSelected As Boolean
    On Error Resume Next
    If Item.Checked = True Then
        subJmlTotal = subJmlTotal + 1
        ReDim Preserve subKdPemeriksa(subJmlTotal)
        subKdPemeriksa(subJmlTotal) = Item.key
    Else
        blnSelected = False
        For i = 1 To subJmlTotal
            If subKdPemeriksa(i) = Item.key Then blnSelected = True
            If blnSelected = True Then
                If i = subJmlTotal Then
                    subKdPemeriksa(i) = ""
                Else
                    subKdPemeriksa(i) = subKdPemeriksa(i + 1)
                End If
            End If
        Next i
        subJmlTotal = subJmlTotal - 1
    End If

    If subJmlTotal = 0 Then
        txtPetugas.BackColor = &HFFFFFF
        chkPerawat.Caption = "Petugas"
    Else
        txtPetugas.BackColor = &HC0FFFF
        chkPerawat.Caption = "Petugas (" & subJmlTotal & " org )"
        Label19.Caption = subJmlTotal
    End If
End Sub

Private Sub subLoadListPemeriksa(Optional strKriteria As String)
    Dim strKey As String

    strSQL = "select * from SIMDaftarPetugas_V " & strKriteria & " order by [Nama Pemeriksa]"
    Call msubRecFO(rs, strSQL)

    With lvPemeriksa
        .ListItems.clear
        For i = 0 To rs.RecordCount - 1
            strKey = "key" & rs(0).Value
            .ListItems.add , strKey, rs(1).Value
            rs.MoveNext
        Next

        .Top = txtPetugas.Top + txtPetugas.Height
        .Left = txtPetugas.Left
        .Height = 1815
        .ColumnHeaders.Item(1).Width = lvPemeriksa.Width - 500

        If subJmlTotal = 0 Then Exit Sub
        For i = 1 To .ListItems.Count
            For j = 1 To subJmlTotal
                If .ListItems(i).key = subKdPemeriksa(j) Then .ListItems(i).Checked = True
            Next j
        Next i
    End With
End Sub

Private Sub lvPemeriksa_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lvPemeriksa.Visible = False
    End If
End Sub

Private Sub txtPetugas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lvPemeriksa.Visible = False
    End If
End Sub

Private Sub txtPetugas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lvPemeriksa.Visible = True Then
            lvPemeriksa.SetFocus
        ElseIf lvPemeriksa.Visible = False Then
            dcStatus2.SetFocus
        End If
    End If
End Sub

Private Sub lvPemeriksa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lvPemeriksa.Visible = False
        dcStatus3.SetFocus
    End If
End Sub

Private Sub subLoadPelayananPerPerawat()
    With fgPerawatPerPelayanan
        For i = 1 To Val(Label19.Caption)       'subJmlTotal
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 1) = Mid(subKdPemeriksa(i), 4, Len(subKdPemeriksa(i)) - 3)
        Next
    End With

    subJmlTotal = 0
    txtPetugas.BackColor = &HFFFFFF
    ReDim Preserve subKdPemeriksa(subJmlTotal)
End Sub

Private Sub subSetGridPerawatPerPelayanan()
    With fgPerawatPerPelayanan
        .Cols = 3
        .Rows = 1

        .MergeCells = flexMergeFree

        .TextMatrix(0, 0) = "Solusi"
        .TextMatrix(0, 1) = "TglMulai"
        .TextMatrix(0, 2) = "IdPegawai"

    End With
End Sub
