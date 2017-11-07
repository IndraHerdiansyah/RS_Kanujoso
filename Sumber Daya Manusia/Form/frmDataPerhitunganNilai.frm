VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDataPerhitunganNilai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Input Penilaian Pegawai"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   ClipControls    =   0   'False
   Icon            =   "frmDataPerhitunganNilai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   8805
   Begin VB.Frame frapegawai 
      Caption         =   "Pegawai Yang di Nilai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   37
      Top             =   1920
      Visible         =   0   'False
      Width           =   6855
      Begin MSDataGridLib.DataGrid dgpegawai 
         Height          =   3495
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraPegawai3 
      Caption         =   "Pegawai Atasan Penilai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   36
      Top             =   3360
      Visible         =   0   'False
      Width           =   6855
      Begin MSDataGridLib.DataGrid dgPegawai3 
         Height          =   3495
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FraPegawai2 
      Caption         =   "Pegawai Penilai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   240
      TabIndex        =   34
      Top             =   2640
      Visible         =   0   'False
      Width           =   6855
      Begin MSDataGridLib.DataGrid dgPegawai2 
         Height          =   3495
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
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
            MarqueeStyle    =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Input Nilai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   44
      Top             =   3600
      Width           =   8775
      Begin VB.Frame Frame5 
         Caption         =   "Nilai"
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
         Left            =   360
         TabIndex        =   47
         Top             =   1320
         Width           =   8055
         Begin VB.TextBox txtKesetiaan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   5
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtPrestasi 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   6
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtTanggungJawab 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   7
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtKetaatan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   8
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtKejujuran 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   9
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox txtKerjasama 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   10
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtPrakarsa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   11
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtKepemimpinan 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6120
            MaxLength       =   3
            TabIndex        =   12
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Kepemimpinan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   55
            Top             =   1800
            Width           =   1680
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Prakarsa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   54
            Top             =   1320
            Width           =   1170
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Kerjasama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   53
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Kejujuran"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4200
            TabIndex        =   52
            Top             =   360
            Width           =   1245
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Ketaatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   51
            Top             =   1800
            Width           =   1170
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Tanggung Jawab"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   50
            Top             =   1320
            Width           =   1875
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Prestasi Kerja"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   49
            Top             =   840
            Width           =   1605
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Nilai Kesetiaan "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   48
            Top             =   360
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Periode"
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
         Left            =   360
         TabIndex        =   45
         Top             =   360
         Width           =   8055
         Begin MSComCtl2.DTPicker dtpBlnHitung 
            Height          =   360
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   108134403
            UpDown          =   -1  'True
            CurrentDate     =   38231
         End
         Begin MSComCtl2.DTPicker dtpAkhir 
            Height          =   360
            Left            =   2880
            TabIndex        =   4
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd MMMM, yyyy"
            Format          =   108134403
            UpDown          =   -1  'True
            CurrentDate     =   38231
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "s / d"
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
            Left            =   2400
            TabIndex        =   46
            Top             =   360
            Width           =   315
         End
      End
   End
   Begin VB.CommandButton cmdCetak2 
      Caption         =   "Cetak Lembar 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   40
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCetak1 
      Caption         =   "Cetak Lembar 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   39
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
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
      Height          =   495
      Left            =   7200
      TabIndex        =   15
      Top             =   7680
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
      Height          =   495
      Left            =   5880
      TabIndex        =   14
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdBaru 
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
      Height          =   495
      Left            =   4560
      TabIndex        =   13
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   20
      Top             =   1080
      Width           =   8775
      Begin VB.TextBox txtNoUrut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7080
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtKdPeg3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3120
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtKdPeg2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtKdPeg1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtJabatanAtasan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   32
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox txtJabatanPenilai 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   16
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtAtasanPenilai 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtPenilai 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtNamaPegawai 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtJK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtJnsPeg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.TextBox txtJabatan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4080
         TabIndex        =   21
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label Label21 
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
         Left            =   4080
         TabIndex        =   33
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label20 
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
         Left            =   4080
         TabIndex        =   31
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Atasan Penilai"
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
         TabIndex        =   30
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Penilai"
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
         TabIndex        =   29
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label2 
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
         Left            =   4080
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin"
         Height          =   210
         Left            =   2160
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label4 
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
         Left            =   4080
         TabIndex        =   24
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   0
      TabIndex        =   17
      Top             =   8640
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox txtTotalHasil 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4440
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   950
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Nilai"
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
         Left            =   4440
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   28
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
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmDataPerhitunganNilai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   6960
      Picture         =   "frmDataPerhitunganNilai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmDataPerhitunganNilai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmDataPerhitunganNilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLSQL As String
Dim strSQL As String
Dim strQuerySQL As String
Dim strLFilterPegawai As String
Dim intLJmlPegawai As Integer
Dim strLIdPegawai As String
Dim strLKdJabatan As String
Dim strLKdPendidikan As String
Const strLOrder As String = "ORDER BY NamaLengkap"
Dim blnLPegawaiFocus As Boolean
Dim intLJmlIndex, intRow As Integer
Dim strKdDetailKomponenIndex As String
Dim strKdKomponenIndex As String

Private Sub cmdBaru_Click()
    Call kosong
    txtNamaPegawai.SetFocus
    frapegawai.Visible = False
    FraPegawai2.Visible = False
    fraPegawai3.Visible = False
    cmdSimpan.Enabled = True
End Sub

Private Sub cmdCetak1_Click()
    frmCetakPenilaianPegawai.Show
End Sub

Private Sub cmdCetak2_Click()
    frmCetakPenilaianPegawaiKeDua.Show
End Sub

Private Sub cmdSimpan_Click()
    On Error GoTo errSimpan

    If Periksa("text", txtNamaPegawai, "Silahkan isi Nama Pegawai") = False Then Exit Sub
    If Periksa("text", txtPenilai, "Silahkan isi Nama Penilai") = False Then Exit Sub
    If Periksa("text", txtAtasanPenilai, "Silahkan isi Atasan Penilai Pegawai") = False Then Exit Sub
    If Periksa("text", txtKesetiaan, "Silahkan isi Nilai Kesetiaan") = False Then Exit Sub
    If Periksa("text", txtPrestasi, "Silahkan isi Nilai Prestasi Kerja") = False Then Exit Sub
    If Periksa("text", txtTanggungJawab, "Silahkan isi Nilai Tanggung Jawab") = False Then Exit Sub
    If Periksa("text", txtKetaatan, "Silahkan isi Nilai Ketaatan") = False Then Exit Sub
    If Periksa("text", txtKejujuran, "Silahkan isi Nilai Kejujuran") = False Then Exit Sub
    If Periksa("text", txtKerjasama, "Silahkan isi Nilai Kerjasama") = False Then Exit Sub
    If Periksa("text", txtPrakarsa, "Silahkan isi Nilai Prakarsa") = False Then Exit Sub
    If Periksa("text", txtKepemimpinan, "Silahkan isi Nilai Kepemimpinan") = False Then Exit Sub

    Dim adoCommand As New ADODB.Command
    With adoCommand
        .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger, adParamReturnValue, , Null)
        If txtNoUrut.Text <> "" Then
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, txtNoUrut.Text)
        Else
            .Parameters.Append .CreateParameter("NoUrut", adChar, adParamInput, 3, Null)
        End If
        .Parameters.Append .CreateParameter("IdPegawai", adChar, adParamInput, 10, Trim(txtKdPeg1.Text))
        .Parameters.Append .CreateParameter("TglAwal", adDate, adParamInput, , Format(dtpBlnHitung.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("TglAkhir", adDate, adParamInput, , Format(dtpAkhir.Value, "yyyy/MM/dd"))
        .Parameters.Append .CreateParameter("NilaiKesetiaan", adVarChar, adParamInput, 15, IIf(txtKesetiaan.Text = "", Null, txtKesetiaan.Text))
        .Parameters.Append .CreateParameter("NilaiPrestasi", adVarChar, adParamInput, 15, IIf(txtPrestasi.Text = "", Null, txtPrestasi.Text))
        .Parameters.Append .CreateParameter("NilaiTanggungJawab", adVarChar, adParamInput, 15, IIf(txtTanggungJawab.Text = "", Null, txtTanggungJawab.Text))
        .Parameters.Append .CreateParameter("NilaiKetaatan", adVarChar, adParamInput, 15, IIf(txtKetaatan.Text = "", Null, txtKetaatan.Text))
        .Parameters.Append .CreateParameter("NilaiKejujuran", adVarChar, adParamInput, 15, IIf(txtKejujuran.Text = "", Null, txtKejujuran.Text))
        .Parameters.Append .CreateParameter("NilaiKerjasama", adVarChar, adParamInput, 15, IIf(txtKerjasama.Text = "", Null, txtKerjasama.Text))
        .Parameters.Append .CreateParameter("NilaiPrakarsa", adVarChar, adParamInput, 15, IIf(txtPrakarsa.Text = "", Null, txtPrakarsa.Text))
        .Parameters.Append .CreateParameter("NilaiKepemimpinan", adVarChar, adParamInput, 15, IIf(txtKepemimpinan.Text = "", Null, txtKepemimpinan.Text))
        .Parameters.Append .CreateParameter("IdPegawai2", adChar, adParamInput, 10, IIf(txtKdPeg2.Text = "", Null, txtKdPeg2.Text))
        .Parameters.Append .CreateParameter("IdPegawai3", adChar, adParamInput, 10, IIf(txtKdPeg3.Text = "", Null, txtKdPeg3.Text))
        .Parameters.Append .CreateParameter("OutputNoUrut", adChar, adParamOutput, 3, Null)

        .ActiveConnection = dbConn
        .CommandText = "AU_PenilaianPegawai"
        .CommandType = adCmdStoredProc
        .Execute

        If Not (.Parameters("RETURN_VALUE").Value = 0) Then
            MsgBox "Ada kesalahan dalam penyimpanan data", vbCritical
            Call deleteADOCommandParameters(adoCommand)
            Set adoCommand = Nothing
            Exit Sub
        Else
            txtNoUrut.Text = .Parameters("OutputNoUrut").Value
            MsgBox "Data telah disimpan..", vbInformation
            cmdSimpan.Enabled = False
        End If

        Call deleteADOCommandParameters(adoCommand)
        Set adoCommand = Nothing

    End With

    Exit Sub
errSimpan:
    Call msubPesanError
    cmdSimpan.Enabled = True
End Sub

Private Sub dgPegawai_DblClick()
    Call dgPegawai_KeyPress(13)
End Sub

Private Sub dgPegawai2_DblClick()
    Call dgPegawai2_KeyPress(13)
End Sub

Private Sub dgPegawai3_DblClick()
    Call dgPegawai3_KeyPress(13)
End Sub

Private Sub dgPegawai_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intLJmlPegawai = 0 Then Exit Sub
        txtNamaPegawai.Text = dgpegawai.Columns(1)
        txtKdPeg1.Text = dgpegawai.Columns(0)
        txtJK.Text = dgpegawai.Columns(2)
        txtJnsPeg.Text = dgpegawai.Columns(3)
        txtJabatan.Text = dgpegawai.Columns(5)
        frapegawai.Visible = False
        txtPenilai.SetFocus
    End If
End Sub

Private Sub dgPegawai2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intLJmlPegawai = 0 Then Exit Sub
        txtPenilai.Text = dgPegawai2.Columns(1)
        txtKdPeg2.Text = dgPegawai2.Columns(0)
        txtJabatanPenilai.Text = dgPegawai2.Columns(5)
        FraPegawai2.Visible = False
        txtAtasanPenilai.SetFocus

    End If
End Sub

Private Sub dgPegawai3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If intLJmlPegawai = 0 Then Exit Sub
        txtAtasanPenilai.Text = dgPegawai3.Columns(1)
        txtKdPeg3.Text = dgPegawai3.Columns(0)
        txtJabatanAtasan.Text = dgPegawai3.Columns(5)
        fraPegawai3.Visible = False
        txtKesetiaan.SetFocus
    End If
End Sub

Private Sub cmdTutup_Click()
    Call frmDaftarPenilaianPegawai.cmdCari_Click
    Unload Me
End Sub

Private Sub dtpAkhir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKesetiaan.SetFocus
End Sub

Private Sub dtpBlnHitung_Change()
    dtpBlnHitung.MaxDate = Now
End Sub

Private Sub dtpBlnHitung_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpAkhir.SetFocus
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    strLSQL = "SELECT * FROM v_S_Pegawai " & strLFilterPegawai
    Call subLoadDataPegawai
    dtpBlnHitung.Value = Format(Now, "dd MMMM, yyyy")
    dtpAkhir.Value = Format(Now, "dd MMMM, yyyy")
End Sub

Private Sub dgPegawai_GotFocus()
    If dgpegawai.Col < 2 Then dgpegawai.Col = 2
End Sub

Private Sub txtAtasanPenilai_Change()
    strLFilterPegawai = "WHERE NamaLengkap LIKE '%" & txtAtasanPenilai.Text & "%'"
    strLSQL = "SELECT * FROM v_S_Pegawai " & strLFilterPegawai
    Call subLoadDataPegawai3
    fraPegawai3.Visible = True
    FraPegawai2.Visible = False
    frapegawai.Visible = False
End Sub

Private Sub txtKejujuran_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKerjasama.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKepemimpinan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKerjasama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPrakarsa.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKesetiaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPrestasi.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtKetaatan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKejujuran.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtNamaPegawai_Change()
    strLFilterPegawai = "WHERE NamaLengkap LIKE '%" & txtNamaPegawai.Text & "%'" '//Yayang.agus 2014-08-08
    'strLFilterPegawai = "WHERE NamaLengkap LIKE '" & txtNamaPegawai.Text & "%'"
    strLSQL = "SELECT * FROM v_S_Pegawai " & strLFilterPegawai
    Call subLoadDataPegawai
    frapegawai.Visible = True
    FraPegawai2.Visible = False
    fraPegawai3.Visible = False
End Sub

Private Sub txtNamaPegawai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If frapegawai.Visible = True Then dgpegawai.SetFocus
End Sub

Private Sub txtNamaPegawai_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
    Select Case KeyAscii
        Case 13
            If frapegawai.Visible = True Then
                dgpegawai.SetFocus
            Else
                txtPenilai.SetFocus
            End If
        Case 27
            frapegawai.Visible = False
    End Select
End Sub

Private Sub txtPenilai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If FraPegawai2.Visible = True Then dgPegawai2.SetFocus
End Sub

Private Sub txtPenilai_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
    Select Case KeyAscii
        Case 13
            If FraPegawai2.Visible = True Then
                dgPegawai2.SetFocus
            Else
                txtAtasanPenilai.SetFocus
            End If
        Case 27
            FraPegawai2.Visible = False
    End Select
End Sub

Private Sub txtAtasanPenilai_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then If fraPegawai3.Visible = True Then dgPegawai3.SetFocus
End Sub

Private Sub txtAtasanPenilai_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
    Select Case KeyAscii
        Case 13
            If fraPegawai3.Visible = True Then
                dgPegawai3.SetFocus
            Else
                dtpBlnHitung.SetFocus
            End If
        Case 27
            fraPegawai3.Visible = False
    End Select
End Sub

Private Sub txtPenilai_Change()

    strLFilterPegawai = "WHERE NamaLengkap LIKE '%" & txtPenilai.Text & "%'"
    strLSQL = "SELECT * FROM v_S_Pegawai " & strLFilterPegawai
    Call subLoadDataPegawai2
    FraPegawai2.Visible = True
    frapegawai.Visible = False
    fraPegawai3.Visible = False
End Sub

Private Sub txtNamaPegawai_GotFocus()
    txtNamaPegawai.SelStart = 0
    txtNamaPegawai.SelLength = Len(txtNamaPegawai.Text)
End Sub

Private Sub subLoadDataPegawai()
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intLJmlPegawai = rs.RecordCount
    With dgpegawai
        Set .DataSource = rs
        .Columns(0).Width = 0
        .Columns(1).Width = 2000
        .Columns(2).Width = 0
        .Columns(3).Width = 2000
        .Columns(4).Width = 0
        .Columns(5).Width = 2000
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(0).Caption = "ID Pegawai"
        .Columns(1).Caption = "Nama Lengkap"
        .Columns(2).Caption = "SEX"
        .Columns(3).Caption = "Jenis Pegawai"
        .Columns(4).Caption = "KD JABATAN"
        .Columns(5).Caption = "Jabatan"
        .Columns(6).Caption = "KD PENDIDIKAN"
        .Columns(7).Caption = "PENDIDIKAN"
    End With
End Sub

Private Sub subLoadDataPegawai2()
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intLJmlPegawai = rs.RecordCount
    With dgPegawai2
        Set .DataSource = rs
        .Columns(0).Width = 0
        .Columns(1).Width = 2000
        .Columns(2).Width = 0
        .Columns(3).Width = 2000
        .Columns(4).Width = 0
        .Columns(5).Width = 2000
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(0).Caption = "ID Pegawai"
        .Columns(1).Caption = "Nama Lengkap"
        .Columns(2).Caption = "SEX"
        .Columns(3).Caption = "Jenis Pegawai"
        .Columns(4).Caption = "KD JABATAN"
        .Columns(5).Caption = "Jabatan"
        .Columns(6).Caption = "KD PENDIDIKAN"
        .Columns(7).Caption = "PENDIDIKAN"
    End With
End Sub

Private Sub subLoadDataPegawai3()
    Set rs = Nothing
    rs.Open strLSQL, dbConn, adOpenForwardOnly, adLockReadOnly
    intLJmlPegawai = rs.RecordCount
    With dgPegawai3
        Set .DataSource = rs
        .Columns(0).Width = 0
        .Columns(1).Width = 2000
        .Columns(2).Width = 0
        .Columns(3).Width = 2000
        .Columns(4).Width = 0
        .Columns(5).Width = 2000
        .Columns(6).Width = 0
        .Columns(7).Width = 0
        .Columns(0).Caption = "ID Pegawai"
        .Columns(1).Caption = "Nama Lengkap"
        .Columns(2).Caption = "SEX"
        .Columns(3).Caption = "Jenis Pegawai"
        .Columns(4).Caption = "KD JABATAN"
        .Columns(5).Caption = "Jabatan"
        .Columns(6).Caption = "KD PENDIDIKAN"
        .Columns(7).Caption = "PENDIDIKAN"
    End With
End Sub

Sub kosong()
    txtNamaPegawai.Text = ""
    txtJK.Text = ""
    txtJnsPeg.Text = ""
    txtJabatan.Text = ""
    txtPenilai.Text = ""
    txtJabatanPenilai.Text = ""
    txtAtasanPenilai.Text = ""
    txtJabatanPenilai.Text = ""
    txtKesetiaan.Text = ""
    txtPrestasi.Text = ""
    txtTanggungJawab.Text = ""
    txtKetaatan.Text = ""
    txtKejujuran.Text = ""
    txtKerjasama.Text = ""
    txtPrakarsa.Text = ""
    txtKepemimpinan.Text = ""
    dtpBlnHitung.Value = Format(Now, "MMMM, yyyy")
    dtpAkhir.Value = Format(Now, "MMMM, yyyy")
    frapegawai.Visible = False
End Sub

Private Sub txtPrakarsa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKepemimpinan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtPrestasi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTanggungJawab.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub

Private Sub txtTanggungJawab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKetaatan.SetFocus
    Call SetKeyPressToNumber(KeyAscii)
End Sub
