VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmMasterDataPenunjang 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Data Master Penunjang"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmMasterDataPenunjang.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   6960
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   375
      Left            =   840
      TabIndex        =   98
      Top             =   7440
      Width           =   1095
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
      Left            =   4440
      TabIndex        =   9
      Top             =   7440
      Width           =   1095
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
      Left            =   3240
      TabIndex        =   8
      Top             =   7440
      Width           =   1095
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
      Left            =   5640
      TabIndex        =   10
      Top             =   7440
      Width           =   1095
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
      Left            =   2040
      TabIndex        =   7
      Top             =   7440
      Width           =   1095
   End
   Begin TabDlg.SSTab sstDataPenunjang 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   9
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Jenis Diklat"
      TabPicture(0)   =   "frmMasterDataPenunjang.frx":0CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label24"
      Tab(0).Control(1)=   "Label23"
      Tab(0).Control(2)=   "Label19(1)"
      Tab(0).Control(3)=   "Label19(4)"
      Tab(0).Control(4)=   "dgJenisDiklat"
      Tab(0).Control(5)=   "txtKdJenisDiklat"
      Tab(0).Control(6)=   "txtJenisDiklat"
      Tab(0).Control(7)=   "txtKdExtJnsDiklat"
      Tab(0).Control(8)=   "txtNmExtJnsDiklat"
      Tab(0).Control(9)=   "chkStsJnsDiklat"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Diklat"
      TabPicture(1)   =   "frmMasterDataPenunjang.frx":0CE6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label27"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label26"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label25"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label19(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label19(6)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "dgDiklat"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "dcJenisDikat"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtKddiklat"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtNamaDiklat"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtKdExtDiklat"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtNmExtDiklat"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkStsDiklat"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Tipe Pekerjaan"
      TabPicture(2)   =   "frmMasterDataPenunjang.frx":0D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkStsTipe"
      Tab(2).Control(1)=   "txtNmExtTipe"
      Tab(2).Control(2)=   "txtKdExtTipe"
      Tab(2).Control(3)=   "txtKdTipe"
      Tab(2).Control(4)=   "txtTipePekerjaan"
      Tab(2).Control(5)=   "dgTipePekerjaan"
      Tab(2).Control(6)=   "Label19(8)"
      Tab(2).Control(7)=   "Label19(7)"
      Tab(2).Control(8)=   "Label28"
      Tab(2).Control(9)=   "Label29"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Eselon"
      TabPicture(3)   =   "frmMasterDataPenunjang.frx":0D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkStsEselon"
      Tab(3).Control(1)=   "txtNmExtEselon"
      Tab(3).Control(2)=   "txtKdExtEselon"
      Tab(3).Control(3)=   "txtKdEselon"
      Tab(3).Control(4)=   "txtNamaEselon"
      Tab(3).Control(5)=   "dgEselon"
      Tab(3).Control(6)=   "Label19(10)"
      Tab(3).Control(7)=   "Label19(9)"
      Tab(3).Control(8)=   "Label21"
      Tab(3).Control(9)=   "Label22"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "Komponen Gaji"
      TabPicture(4)   =   "frmMasterDataPenunjang.frx":0D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label11"
      Tab(4).Control(1)=   "Label10"
      Tab(4).Control(2)=   "Label19(11)"
      Tab(4).Control(3)=   "Label19(12)"
      Tab(4).Control(4)=   "dgKomponenGaji"
      Tab(4).Control(5)=   "txtKodeKelompok"
      Tab(4).Control(6)=   "txtKelompokGaji"
      Tab(4).Control(7)=   "txtKdExtGaji"
      Tab(4).Control(8)=   "txtNmExtGaji"
      Tab(4).Control(9)=   "chkStsGaji"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Jenis Hukuman"
      TabPicture(5)   =   "frmMasterDataPenunjang.frx":0D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chkStsHukuman"
      Tab(5).Control(1)=   "txtNmExtHukuman"
      Tab(5).Control(2)=   "txtKdExtHukuman"
      Tab(5).Control(3)=   "txtKdJenisHukum"
      Tab(5).Control(4)=   "txtJenisHukum"
      Tab(5).Control(5)=   "dgJenisHukum"
      Tab(5).Control(6)=   "Label19(3)"
      Tab(5).Control(7)=   "Label19(2)"
      Tab(5).Control(8)=   "Label19(0)"
      Tab(5).Control(9)=   "Label20"
      Tab(5).ControlCount=   10
      TabCaption(6)   =   "Status Perkawinan"
      TabPicture(6)   =   "frmMasterDataPenunjang.frx":0D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label1"
      Tab(6).Control(1)=   "Label2"
      Tab(6).Control(2)=   "Label19(13)"
      Tab(6).Control(3)=   "Label19(14)"
      Tab(6).Control(4)=   "dgStatusPerkawinan"
      Tab(6).Control(5)=   "txtKdStatusPerkawinan"
      Tab(6).Control(6)=   "txtStatusPerkawinan"
      Tab(6).Control(7)=   "txtKdExtKawin"
      Tab(6).Control(8)=   "txtNmExtKawin"
      Tab(6).Control(9)=   "chkStsKawin"
      Tab(6).ControlCount=   10
      TabCaption(7)   =   "Agama"
      TabPicture(7)   =   "frmMasterDataPenunjang.frx":0D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label3"
      Tab(7).Control(1)=   "Label4"
      Tab(7).Control(2)=   "Label19(15)"
      Tab(7).Control(3)=   "Label19(16)"
      Tab(7).Control(4)=   "dgAgama"
      Tab(7).Control(5)=   "txtAgama"
      Tab(7).Control(6)=   "txtKdAgama"
      Tab(7).Control(7)=   "txtKdExtAgama"
      Tab(7).Control(8)=   "txtNmExtAgama"
      Tab(7).Control(9)=   "chkStsAgama"
      Tab(7).ControlCount=   10
      TabCaption(8)   =   "Golongan Darah"
      TabPicture(8)   =   "frmMasterDataPenunjang.frx":0DAA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "chkSts"
      Tab(8).Control(1)=   "txtNmExt"
      Tab(8).Control(2)=   "txtKdExt"
      Tab(8).Control(3)=   "txtGolonganDarah"
      Tab(8).Control(4)=   "txtKdGolonganDarah"
      Tab(8).Control(5)=   "dgGolonganDarah"
      Tab(8).Control(6)=   "Label19(18)"
      Tab(8).Control(7)=   "Label19(17)"
      Tab(8).Control(8)=   "Label6"
      Tab(8).Control(9)=   "Label5"
      Tab(8).ControlCount=   10
      Begin VB.CheckBox chkSts 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   57
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         MaxLength       =   50
         TabIndex        =   56
         Top             =   2280
         Width           =   4695
      End
      Begin VB.TextBox txtKdExt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73320
         MaxLength       =   15
         TabIndex        =   55
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsAgama 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   51
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtAgama 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   50
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox txtKdExtAgama 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         MaxLength       =   15
         TabIndex        =   49
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsKawin 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   45
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtKawin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   44
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox txtKdExtKawin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73080
         MaxLength       =   15
         TabIndex        =   43
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsGaji 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   34
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtGaji 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txtKdExtGaji 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   15
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsEselon 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   28
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtEselon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox txtKdExtEselon 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         MaxLength       =   15
         TabIndex        =   26
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsTipe 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   22
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtTipe 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   21
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txtKdExtTipe 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   15
         TabIndex        =   20
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsDiklat 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5160
         TabIndex        =   16
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtDiklat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   15
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox txtKdExtDiklat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   14
         Top             =   2280
         Width           =   1815
      End
      Begin VB.CheckBox chkStsJnsDiklat 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   5
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtJnsDiklat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox txtKdExtJnsDiklat 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73560
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.CheckBox chkStsHukuman 
         Alignment       =   1  'Right Justify
         Caption         =   "Status Aktif"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69840
         TabIndex        =   39
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtNmExtHukuman 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   38
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txtKdExtHukuman 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -73440
         MaxLength       =   15
         TabIndex        =   37
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox txtGolonganDarah 
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
         Left            =   -73320
         MaxLength       =   2
         TabIndex        =   54
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtKdGolonganDarah 
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
         Left            =   -73320
         MaxLength       =   2
         TabIndex        =   53
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtKdAgama 
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
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   47
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtAgama 
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
         Left            =   -73560
         MaxLength       =   20
         TabIndex        =   48
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txtStatusPerkawinan 
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
         Left            =   -73080
         MaxLength       =   20
         TabIndex        =   42
         Top             =   1560
         Width           =   4455
      End
      Begin VB.TextBox txtKdStatusPerkawinan 
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
         Left            =   -73080
         MaxLength       =   10
         TabIndex        =   41
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtKdJenisHukum 
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
         Left            =   -73440
         MaxLength       =   3
         TabIndex        =   35
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtJenisHukum 
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   36
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtKelompokGaji 
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtKodeKelompok 
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
         Left            =   -73440
         MaxLength       =   2
         TabIndex        =   30
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtKdEselon 
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
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   24
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNamaEselon 
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
         Left            =   -73560
         MaxLength       =   20
         TabIndex        =   25
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txtKdTipe 
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
         Left            =   -73440
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtTipePekerjaan 
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   19
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox txtNamaDiklat 
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
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txtKddiklat 
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
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtJenisDiklat 
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
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txtKdJenisDiklat 
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
         Left            =   -73560
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1200
         Width           =   735
      End
      Begin MSDataGridLib.DataGrid dgJenisDiklat 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   6
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcJenisDikat 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
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
      Begin MSDataGridLib.DataGrid dgEselon 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   29
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgKomponenGaji 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   69
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgTipePekerjaan 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   23
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgJenisHukum 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   40
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgStatusPerkawinan 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   46
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgAgama 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   52
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgGolonganDarah 
         Height          =   3135
         Left            =   -74760
         TabIndex        =   58
         Top             =   2760
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   5530
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgDiklat 
         Height          =   2775
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4895
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
         Index           =   18
         Left            =   -74760
         TabIndex        =   97
         Top             =   2295
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   96
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   16
         Left            =   -74745
         TabIndex        =   95
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   15
         Left            =   -74760
         TabIndex        =   94
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   14
         Left            =   -74760
         TabIndex        =   93
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   13
         Left            =   -74760
         TabIndex        =   92
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   12
         Left            =   -74760
         TabIndex        =   91
         Top             =   2280
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   11
         Left            =   -74760
         TabIndex        =   90
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   10
         Left            =   -74745
         TabIndex        =   89
         Top             =   2295
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   9
         Left            =   -74760
         TabIndex        =   88
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   8
         Left            =   -74760
         TabIndex        =   87
         Top             =   2295
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   7
         Left            =   -74760
         TabIndex        =   86
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   6
         Left            =   255
         TabIndex        =   85
         Top             =   2640
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   84
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   4
         Left            =   -74745
         TabIndex        =   83
         Top             =   2295
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   82
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Nama External"
         Height          =   195
         Index           =   3
         Left            =   -74745
         TabIndex        =   81
         Top             =   2295
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode External"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   80
         Top             =   1920
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Golongan Darah"
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
         Left            =   -74760
         TabIndex        =   79
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   78
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   77
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agama"
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
         Left            =   -74760
         TabIndex        =   76
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status Perkawinan"
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
         Left            =   -74760
         TabIndex        =   75
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   74
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   73
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Hukuman"
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
         Left            =   -74760
         TabIndex        =   72
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Komponen Gaji"
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
         Left            =   -74760
         TabIndex        =   71
         Top             =   1560
         Width           =   1230
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   70
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   68
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Nama Eselon"
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
         Left            =   -74760
         TabIndex        =   67
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Kode "
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
         Left            =   -74760
         TabIndex        =   66
         Top             =   1200
         Width           =   480
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Tipe Pekerjaan"
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
         Left            =   -74760
         TabIndex        =   65
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         TabIndex        =   64
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Nama Diklat"
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
         TabIndex        =   63
         Top             =   1560
         Width           =   945
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Diklat"
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
         TabIndex        =   62
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Diklat"
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
         Left            =   -74760
         TabIndex        =   61
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Kode"
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
         Left            =   -74760
         TabIndex        =   60
         Top             =   1200
         Width           =   420
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   59
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
      Left            =   5160
      Picture         =   "frmMasterDataPenunjang.frx":0DC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmMasterDataPenunjang.frx":1B4E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMasterDataPenunjang.frx":450F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmMasterDataPenunjang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub subDCSource()
    strSQL = "SELECT * FROM JenisDiklat Where StatusEnabled='1' order by JenisDiklat"
    Call msubDcSource(dcJenisDikat, rs, strSQL)
End Sub

Sub sp_simpan(f_Status As String)
    Select Case sstDataPenunjang.Tab
        Case 0 ' Jenis diklat
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisDiklat", adVarChar, adParamInput, 2, Trim(txtKdJenisDiklat))
                .Parameters.Append .CreateParameter("JenisDiklat", adVarChar, adParamInput, 50, Trim(txtJenisDiklat))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtJnsDiklat.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtJnsDiklat.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsJnsDiklat.Value)
                .Parameters.Append .CreateParameter("OutputKdJenisDiklat", adChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_JenisDiklat"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                    If Not IsNull(.Parameters("OutputKdJenisDiklat").Value) Then txtKdJenisDiklat = .Parameters("OutputKdJenisDiklat").Value
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

        Case 1 ' diklat
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdDiklat", adVarChar, adParamInput, 5, IIf(txtKddiklat.Text = "", Null, Trim(txtKddiklat.Text)))
                .Parameters.Append .CreateParameter("NamaDiklat", adVarChar, adParamInput, 100, Trim(txtNamaDiklat))
                .Parameters.Append .CreateParameter("KdJenisDiklat", adVarChar, adParamInput, 2, Trim(dcJenisDikat.BoundText))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtDiklat.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtDiklat.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsDiklat.Value)
                .Parameters.Append .CreateParameter("OutputKdDiklat", adVarChar, adParamOutput, 5, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_Diklat"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                    If Not IsNull(.Parameters("OutputKdDiklat").Value) Then txtKddiklat = .Parameters("OutputKdDiklat").Value
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click
        Case 2 ' tipe pekerjaan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdTipe", adVarChar, adParamInput, 2, Trim(txtKdTipe))
                .Parameters.Append .CreateParameter("TipePekerjaan", adVarChar, adParamInput, 50, Trim(txtTipePekerjaan))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtTipe.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtTipe.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsTipe.Value)
                .Parameters.Append .CreateParameter("OutputKdTipe", adChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_TipePekerjaan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                    If Not IsNull(.Parameters("OutputKdTipe").Value) Then txtKdTipe = .Parameters("OutputKdTipe").Value
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click
        Case 3 ' esselon
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdEselon", adVarChar, adParamInput, 2, Trim(txtKdEselon))
                .Parameters.Append .CreateParameter("NamaEselon", adVarChar, adParamInput, 20, Trim(txtNamaEselon))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtEselon.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtEselon.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsEselon.Value)
                .Parameters.Append .CreateParameter("OutputKdEselon", adChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_Eselon"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                    If Not IsNull(.Parameters("OutputKdEselon").Value) Then txtKdEselon = .Parameters("OutputKdEselon").Value
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

        Case 4 'Komponen Gaji
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdKomponenGaji", adChar, adParamInput, 2, Trim(txtKodeKelompok))
                .Parameters.Append .CreateParameter("KomponenGaji", adVarChar, adParamInput, 50, Trim(txtKelompokGaji))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtGaji.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtGaji.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsGaji.Value)
                .Parameters.Append .CreateParameter("OutputKdKomponenGaji", adChar, adParamOutput, 2, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_KomponenGaji"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                    If Not IsNull(.Parameters("OutputKdKomponenGaji").Value) Then txtKodeKelompok = .Parameters("OutputKdKomponenGaji").Value
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

        Case 5 ' jenis hukum
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdJenisHukuman", adVarChar, adParamInput, 3, Trim(txtKdJenisHukum))
                .Parameters.Append .CreateParameter("JenisHukuman", adVarChar, adParamInput, 50, Trim(txtJenisHukum))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtHukuman.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtHukuman.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsHukuman.Value)
                .Parameters.Append .CreateParameter("OutputKdJenisHukuman", adChar, adParamOutput, 3, Null)

                .ActiveConnection = dbConn
                .CommandText = "AU_JenisHukuman"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                    If Not IsNull(.Parameters("OutputKdJenisHukuman").Value) Then txtKdJenisHukum = .Parameters("OutputKdJenisHukuman").Value
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click
        Case 6 ' status perkawinan
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdStatusPerkawinan", adChar, adParamInput, 10, Trim(txtKdStatusPerkawinan))
                .Parameters.Append .CreateParameter("StatusPerkawinan", adVarChar, adParamInput, 20, Trim(txtStatusPerkawinan))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtKawin.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtKawin.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsKawin.Value)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_StatusPerkawinan"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

        Case 7 ' Agama
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdAgama", adChar, adParamInput, 2, Trim(txtKdAgama))
                .Parameters.Append .CreateParameter("Agama", adVarChar, adParamInput, 20, Trim(txtAgama))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExtAgama.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExtAgama.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkStsAgama.Value)
                .Parameters.Append .CreateParameter("Status", adChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_Agama"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

        Case 8 ' Golongan darah
            Set dbcmd = New ADODB.Command
            With dbcmd
                .Parameters.Append .CreateParameter("return_value", adInteger, adParamReturnValue, , Null)
                .Parameters.Append .CreateParameter("KdGolonganDarah", adChar, adParamInput, 2, Trim(txtKdGolonganDarah))
                .Parameters.Append .CreateParameter("GolonganDarah", adVarChar, adParamInput, 2, Trim(txtGolonganDarah))
                .Parameters.Append .CreateParameter("KodeExternal", adVarChar, adParamInput, 15, Trim(txtKdExt.Text))
                .Parameters.Append .CreateParameter("NamaExternal", adVarChar, adParamInput, 50, Trim(txtNmExt.Text))
                .Parameters.Append .CreateParameter("StatusEnabled", adTinyInt, adParamInput, , chkSts.Value)
                .Parameters.Append .CreateParameter("Status", adVarChar, adParamInput, 1, f_Status)

                .ActiveConnection = dbConn
                .CommandText = "AUD_GolonganDarah"
                .CommandType = adCmdStoredProc
                .Execute

                If Not (.Parameters("return_value").Value = 0) Then
                    MsgBox "Ada kesalahan dalam pemasukan data!", vbExclamation, "Validasi"

                Else
                End If
                Call deleteADOCommandParameters(dbcmd)

            End With
            cmdBatal_Click

    End Select
End Sub

Private Sub cmdBatal_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 ' Jenis diklat
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 1 ' diklat
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 2 ' tipe pekerjaan
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 3 ' ESELON
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 4 'Komponen Gaji
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 5 ' Jenis Hukum
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 6 ' Status perkawinan
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 7 ' agama
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
        Case 8 ' gol darah
            Call subKosong
            cmdHapus.Enabled = True
            cmdSimpan.Enabled = True
    End Select
End Sub

Private Sub cmdCetak_Click()
    On Error GoTo hell
    Select Case sstDataPenunjang.Tab
        Case 0
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmjenisdiklat.Show
        Case 1
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmdiklat.Show
        Case 2
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmtipepekerjaan.Show
        Case 3
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmeselon.Show
        Case 4
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmcetakkomponengaji.Show
        Case 5
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmjenishukuman.Show
        Case 6
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmstatusperkawinan.Show
        Case 7
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmagama.Show
        Case 8
            vLaporan = ""
            If MsgBox("Apakah Anda Ingin Langsung Mencetak Laporan?" & vbNewLine & "Pilih No Jika Ingin Ditampilkan Terlebih Dahulu", vbYesNo, "Medifirst2000 - Cetak Laporan") = vbNo Then vLaporan = "view"
            frmgolongandarah.Show
    End Select
hell:
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Hapus Data ini? ", vbQuestion + vbYesNo, "Konfirmasi") = vbNo Then Exit Sub
    Select Case sstDataPenunjang.Tab
        Case 0 'Jenis diklat
'            Set rs = Nothing
'            If txtKdJenisDiklat.Text = "" Then Exit Sub
'            strSQL = "delete JenisDiklat  where KdJenisDiklat= '" & txtKdJenisDiklat & "'"
'            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
'            Set rs = Nothing

            If Periksa("text", txtJenisDiklat, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from Diklat where KdJenisDiklat='" & txtKdJenisDiklat & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                dbConn.Execute "Delete JenisDiklat WHERE KdJenisDiklat = '" & txtKdJenisDiklat.Text & "'"
            End If
            
        Case 1 ' diklat
'            Set rs = Nothing
'            If txtKddiklat.Text = "" Then Exit Sub
'            strSQL = "delete Diklat  where KdDiklat= '" & txtKddiklat & "'"
'            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
'            Set rs = Nothing

            If Periksa("text", txtNamaDiklat, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from RiwayatDiklatPelatihan where KdDiklat='" & txtKddiklat & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                dbConn.Execute "Delete Diklat WHERE KdDiklat = '" & txtKddiklat.Text & "'"
            End If
            
        Case 2 ' tipe pekerjaan
            Set rs = Nothing
            If txtKdTipe.Text = "" Then Exit Sub
            strSQL = "delete TipePekerjaan where KdTipe= '" & txtKdTipe & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
            
        Case 3 'eselon
'            Set rs = Nothing
'            If txtKdEselon.Text = "" Then Exit Sub
'            strSQL = "delete Eselon where KdEselon= '" & txtKdEselon & "'"
'            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
'            Set rs = Nothing

            If Periksa("text", txtNamaEselon, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from RiwayatTempatBertugas where KdEselon='" & txtKdEselon & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                dbConn.Execute "Delete Eselon WHERE KdEselon = '" & txtKdEselon.Text & "'"
            End If
            
        Case 4 'Komponen Gaji
'            Set rs = Nothing
'            If txtKodeKelompok.Text = "" Then Exit Sub
'            strSQL = "delete komponengaji where kdkomponengaji = '" & txtKodeKelompok & "'"
'            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
'            Set rs = Nothing

            If Periksa("text", txtKelompokGaji, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from DetailRiwayatGajiPegawai where KdKomponenGaji='" & txtKodeKelompok & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                dbConn.Execute "Delete KomponenGaji WHERE KdKomponenGaji = '" & txtKodeKelompok.Text & "'"
            End If

        Case 5 'Jenis hukuman
            Set rs = Nothing
            If txtKdJenisHukum.Text = "" Then Exit Sub
            strSQL = "delete JenisHukuman where KdJenisHukuman = '" & txtKdJenisHukum & "'"
            rs.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
            Set rs = Nothing
            
        Case 6 'status perkwianan
'            Set rs = Nothing
'            If txtKdStatusPerkawinan.Text = "" Then Exit Sub
'            Call sp_simpan("D")
'            Set rs = Nothing

            If Periksa("text", txtStatusPerkawinan, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from DataCurrentPegawai where KdStatusPerkawinan='" & txtKdStatusPerkawinan & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                Set rs = Nothing
                If txtKdStatusPerkawinan.Text = "" Then Exit Sub
                Call sp_simpan("D")
                Set rs = Nothing
            End If

        Case 7 'Agama
'            Set rs = Nothing
'            Call sp_simpan("D")
'            Set rs = Nothing

            If Periksa("text", txtAgama, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from DataCurrentPegawai where KdAgama='" & txtKdAgama & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                Set rs = Nothing
                Call sp_simpan("D")
                Set rs = Nothing
                End If
            
        Case 8 'golongan darah
'            If txtKdGolonganDarah.Text = "" Then Exit Sub
'            Set rs = Nothing
'            Call sp_simpan("")
'            Set rs = Nothing

            If Periksa("text", txtGolonganDarah, "Pilih Data yang akan dihapus") = False Then Exit Sub
            Set rs = Nothing
            strSQL = "Select * from DataCurrentPegawai where KdGolonganDarah='" & txtKdGolonganDarah & "'"
            Call msubRecFO(rs, strSQL)
            If rs.EOF = False Then
                MsgBox "data tidak bisa dihapus, data masih digunakan di tabel lain", vbExclamation, "Validasi"
                Exit Sub
            Else
                If txtKdGolonganDarah.Text = "" Then Exit Sub
                Set rs = Nothing
                Call sp_simpan("")
                Set rs = Nothing
                End If
    End Select
    MsgBox "Data berhasil dihapus..", vbInformation, "Informasi"
    Call subLoadGridSource
    Call subKosong
End Sub

Private Sub cmdSimpan_Click()
    Select Case sstDataPenunjang.Tab
        Case 0 'Jenis diklat
            If Periksa("text", txtJenisDiklat, "Isi Jenis diklat !") = False Then Exit Sub
            Call sp_simpan("")
        Case 1 'diklat
            If Periksa("text", txtNamaDiklat, "Isi Jenis nama diklat !") = False Then Exit Sub
            If Periksa("datacombo", dcJenisDikat, "Isi Jenisdiklat !") = False Then Exit Sub
            Call sp_simpan("")
        Case 2 'tipe pekerjaan
            If Periksa("text", txtTipePekerjaan, "Isi tipe pekerjaan!") = False Then Exit Sub
            Call sp_simpan("")
        Case 3 'Eselon
            If Periksa("text", txtNamaEselon, "Isi Nama eselon!") = False Then Exit Sub
            Call sp_simpan("")
        Case 4 'komponen gaji
            If Periksa("text", txtKelompokGaji, "Isi Kelompok gaji") = False Then Exit Sub
            Call sp_simpan("")
        Case 5 'Jenis hukum
            If Periksa("text", txtJenisHukum, "Isi Jenis hukum!") = False Then Exit Sub
            Call sp_simpan("")
        Case 6 'status perkawinan
            If Periksa("text", txtStatusPerkawinan, "Isi Status Perkawinan!") = False Then Exit Sub
            Call sp_simpan("A")
        Case 7 'agama
            If Periksa("text", txtAgama, "Isi Agama!") = False Then Exit Sub
            Call sp_simpan("A")
        Case 8 'Golongan darah
            If Periksa("text", txtGolonganDarah, "Isi Golongan darah!") = False Then Exit Sub
            Call sp_simpan("A")
    End Select
    MsgBox "Data berhasil disimpan..", vbInformation, "Informasi"
    Call subLoadGridSource
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcJenisDikat_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then txtKdExtDiklat.SetFocus
On Error GoTo Errload
If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then
        If Len(Trim(dcJenisDikat.Text)) = 0 Then txtKdExtDiklat.SetFocus: Exit Sub
        If dcJenisDikat.MatchedWithList = True Then txtKdExtDiklat.SetFocus: Exit Sub
        Call msubRecFO(dbRst, "SELECT KdJenisDiklat,JenisDiklat FROM JenisDiklat WHERE JenisDiklat LIKE '%" & dcJenisDikat.Text & "%'")
        If dbRst.EOF = True Then Exit Sub
        dcJenisDikat.BoundText = dbRst(0).Value
        dcJenisDikat.Text = dbRst(1).Value
    End If
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub dgAgama_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgAgama.ApproxCount = 0 Then Exit Sub
    txtKdAgama.Text = dgAgama.Columns(0).Value
    txtAgama.Text = dgAgama.Columns(1).Value

    txtKdExtAgama.Text = dgAgama.Columns(2)
    txtNmExtAgama.Text = dgAgama.Columns(3)

    If dgAgama.Columns(4).Value = "<Type mismacth>" Then
        chkStsAgama.Value = 0
    Else
        If dgAgama.Columns(4).Value = 1 Then
            chkStsAgama.Value = 1
        Else
            chkStsAgama.Value = 0
        End If
    End If

End Sub

Private Sub dgDiklat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKddiklat.Text = dgDiklat.Columns(0).Value
    txtNamaDiklat.Text = dgDiklat.Columns(1).Value
    dcJenisDikat.Text = dgDiklat.Columns(2).Value

    txtKdExtDiklat.Text = dgDiklat.Columns(3)
    txtNmExtDiklat.Text = dgDiklat.Columns(4)

    If dgDiklat.Columns(5).Value = "<Type mismacth>" Then
        chkStsDiklat.Value = 0
    Else
        If dgDiklat.Columns(5).Value = 1 Then
            chkStsDiklat.Value = 1
        Else
            chkStsDiklat.Value = 0
        End If
    End If

End Sub

Private Sub dgEselon_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdEselon.Text = dgEselon.Columns(0).Value
    txtNamaEselon.Text = dgEselon.Columns(1).Value

    txtKdExtEselon.Text = dgEselon.Columns(2)
    txtNmExtEselon.Text = dgEselon.Columns(3)

    If dgEselon.Columns(4).Value = "<Type mismacth>" Then
        chkStsEselon.Value = 0
    Else
        If dgEselon.Columns(4).Value = 1 Then
            chkStsEselon.Value = 1
        Else
            chkStsEselon.Value = 0
        End If
    End If

End Sub

Private Sub dgGolonganDarah_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgGolonganDarah.ApproxCount = 0 Then Exit Sub
    txtKdGolonganDarah.Text = dgGolonganDarah.Columns(0).Value
    txtGolonganDarah.Text = dgGolonganDarah.Columns(1).Value

    txtKdExt.Text = dgGolonganDarah.Columns(2)
    txtNmExt.Text = dgGolonganDarah.Columns(3)

    If dgGolonganDarah.Columns(4).Value = "<Type mismacth>" Then
        chkSts.Value = 0
    Else
        If dgGolonganDarah.Columns(4).Value = 1 Then
            chkSts.Value = 1
        Else
            chkSts.Value = 0
        End If
    End If

End Sub

Private Sub dgJenisDiklat_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJenisDiklat.Text = dgJenisDiklat.Columns(0).Value
    txtJenisDiklat.Text = dgJenisDiklat.Columns(1).Value

    txtKdExtJnsDiklat.Text = dgJenisDiklat.Columns(2)
    txtNmExtJnsDiklat.Text = dgJenisDiklat.Columns(3)

    If dgJenisDiklat.Columns(4).Value = "<Type mismacth>" Then
        chkStsJnsDiklat.Value = 0
    Else
        If dgJenisDiklat.Columns(4).Value = 1 Then
            chkStsJnsDiklat.Value = 1
        Else
            chkStsJnsDiklat.Value = 0
        End If
    End If

End Sub

Private Sub dgJenisHukum_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdJenisHukum.Text = dgJenisHukum.Columns(0).Value
    txtJenisHukum.Text = dgJenisHukum.Columns(1).Value

    txtKdExtHukuman.Text = dgJenisHukum.Columns(2)
    txtNmExtHukuman.Text = dgJenisHukum.Columns(3)

    If dgJenisHukum.Columns(4).Value = "<Type mismacth>" Then
        chkStsHukuman.Value = 0
    Else
        If dgJenisHukum.Columns(4).Value = 1 Then
            chkStsHukuman.Value = 1
        Else
            chkStsHukuman.Value = 0
        End If
    End If
End Sub

Private Sub dgKomponenGaji_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKodeKelompok.Text = dgKomponenGaji.Columns(0).Value
    txtKelompokGaji.Text = dgKomponenGaji.Columns(1).Value

    txtKdExtGaji.Text = dgKomponenGaji.Columns(2)
    txtNmExtGaji.Text = dgKomponenGaji.Columns(3)

    If dgKomponenGaji.Columns(4).Value = "<Type mismacth>" Then
        chkStsGaji.Value = 0
    Else
        If dgKomponenGaji.Columns(4).Value = 1 Then
            chkStsGaji.Value = 1
        Else
            chkStsGaji.Value = 0
        End If
    End If

    txtKodeKelompok.Enabled = False
End Sub

Private Sub dgStatusPerkawinan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    If dgStatusPerkawinan.ApproxCount = 0 Then Exit Sub
    txtKdStatusPerkawinan.Text = dgStatusPerkawinan.Columns(0).Value
    txtStatusPerkawinan.Text = dgStatusPerkawinan.Columns(1).Value

    txtKdExtKawin.Text = dgStatusPerkawinan.Columns(3)
    txtNmExtKawin.Text = dgStatusPerkawinan.Columns(4)

    If dgStatusPerkawinan.Columns(5).Value = "<Type mismacth>" Then
        chkStsKawin.Value = 0
    Else
        If dgStatusPerkawinan.Columns(5).Value = 1 Then
            chkStsKawin.Value = 1
        Else
            chkStsKawin.Value = 0
        End If
    End If
End Sub

Private Sub dgTipePekerjaan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    txtKdTipe.Text = dgTipePekerjaan.Columns(0).Value
    txtTipePekerjaan.Text = dgTipePekerjaan.Columns(1).Value

    txtKdExtTipe.Text = dgTipePekerjaan.Columns(2)
    txtNmExtTipe.Text = dgTipePekerjaan.Columns(3)

    If dgTipePekerjaan.Columns(4).Value = "<Type mismacth>" Then
        chkStsTipe.Value = 0
    Else
        If dgTipePekerjaan.Columns(4).Value = 1 Then
            chkStsTipe.Value = 1
        Else
            chkStsTipe.Value = 0
        End If
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    Call subDCSource
    sstDataPenunjang.Tab = 0
    Call subLoadGridSource
End Sub

Sub subKosong()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab
        Case 0 'jenis diklat
            txtKdJenisDiklat.Text = ""
            txtJenisDiklat.Text = ""
            txtKdExtJnsDiklat = ""
            txtNmExtJnsDiklat = ""
            chkStsJnsDiklat.Value = 1
            txtJenisDiklat.SetFocus
        Case 1 'diklat
            txtKddiklat.Text = ""
            txtNamaDiklat.Text = ""
            dcJenisDikat.Text = ""
            txtKdExtDiklat = ""
            txtNmExtDiklat = ""
            chkStsDiklat.Value = 1
            dcJenisDikat.SetFocus
        Case 2 'tipe pekerjaan
            txtKdTipe.Text = ""
            txtTipePekerjaan.Text = ""
            txtKdExtTipe = ""
            txtNmExtTipe = ""
            chkStsTipe.Value = 1
            txtTipePekerjaan.SetFocus
        Case 3 'Eselon
            txtKdEselon.Text = ""
            txtNamaEselon.Text = ""
            txtKdExtEselon = ""
            txtNmExtEselon = ""
            chkStsEselon.Value = 1
            txtNamaEselon.SetFocus
        Case 4 'Komponen Gaji
            txtKodeKelompok.Text = ""
            txtKelompokGaji.Text = ""
            txtKdExtGaji = ""
            txtNmExtGaji = ""
            chkStsGaji.Value = 1
            txtKelompokGaji.SetFocus
        Case 5 'Jenis hukum
            txtKdJenisHukum.Text = ""
            txtJenisHukum.Text = ""
            txtKdExtHukuman = ""
            txtNmExtHukuman = ""
            chkStsHukuman.Value = 1
            txtJenisHukum.SetFocus
        Case 6 'status perkawinan
            txtKdStatusPerkawinan.Text = ""
            txtStatusPerkawinan.Text = ""
            txtKdExtKawin = ""
            txtNmExtKawin = ""
            chkStsKawin.Value = 1
            txtStatusPerkawinan.SetFocus
        Case 7 'agama
            txtKdAgama.Text = ""
            txtAgama.Text = ""
            txtKdExtAgama = ""
            txtNmExtAgama = ""
            chkStsAgama.Value = 1
            txtAgama.SetFocus
        Case 8 'golongan darah
            txtKdGolonganDarah.Text = ""
            txtGolonganDarah.Text = ""
            txtKdExt = ""
            txtNmExt = ""
            chkSts.Value = 1
            txtGolonganDarah.SetFocus

    End Select
End Sub

Private Sub sstDataPenunjang_Click(PreviousTab As Integer)
    Call subDCSource
    Call subLoadGridSource
End Sub

Private Sub txtAgama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtAgama.SetFocus
End Sub

Private Sub txtAgama_LostFocus()
    txtAgama.Text = StrConv(txtAgama, vbProperCase)
End Sub

Private Sub txtGolonganDarah_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExt.SetFocus
End Sub

Private Sub txtGolonganDarah_LostFocus()
    txtGolonganDarah.Text = StrConv(txtGolonganDarah, vbUpperCase)
End Sub

Private Sub txtJenisDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtJnsDiklat.SetFocus
End Sub

Private Sub txtJenisHukum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtHukuman.SetFocus
End Sub

Private Sub txtKdExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExt.SetFocus
End Sub

Private Sub txtKdExtAgama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtAgama.SetFocus
End Sub

Private Sub txtKdExtDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtDiklat.SetFocus
End Sub

Private Sub txtKdExtEselon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtEselon.SetFocus
End Sub

Private Sub txtKdExtGaji_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtGaji.SetFocus
End Sub

Private Sub txtKdExtHukuman_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtHukuman.SetFocus
End Sub

Private Sub txtKdExtKawin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtKawin.SetFocus
End Sub

Private Sub txtKdExtTipe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtTipe.SetFocus
End Sub

Private Sub txtKelompokGaji_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtGaji.SetFocus
End Sub

Private Sub txtNamaDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dcJenisDikat.SetFocus
End Sub

Private Sub txtNamaEselon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtEselon.SetFocus
End Sub

Private Sub txtKodeKelompok_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKelompokGaji.SetFocus
End Sub

Sub subLoadGridSource()
    On Error Resume Next
    Select Case sstDataPenunjang.Tab

        Case 0 'jenis diklat
            Set rs = Nothing
            strSQL = "select * from JenisDiklat"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisDiklat.DataSource = rs
            dgJenisDiklat.Columns(0).Width = 750
            dgJenisDiklat.Columns(0).Alignment = vbCenter
            dgJenisDiklat.Columns(1).Width = 4200
            dgJenisDiklat.ReBind
            Set rs = Nothing

        Case 1 ' diklat
            Set rs = Nothing
            strSQL = "SELECT     dbo.Diklat.KdDiklat, dbo.Diklat.NamaDiklat, dbo.JenisDiklat.JenisDiklat, " & _
            " dbo.Diklat.KodeExternal, dbo.Diklat.NamaExternal, dbo.Diklat.StatusEnabled FROM dbo.Diklat INNER JOIN " & _
            " dbo.JenisDiklat ON dbo.Diklat.KdJenisDiklat = dbo.JenisDiklat.KdJenisDiklat"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgDiklat.DataSource = rs
            dgDiklat.Columns(0).Width = 750
            dgDiklat.Columns(0).Alignment = vbCenter
            dgDiklat.Columns(1).Width = 3000
            dgDiklat.ReBind
            Set rs = Nothing

        Case 2 ' Tipe  pekerjaan
            Set rs = Nothing
            strSQL = "SELECT * From TipePekerjaan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgTipePekerjaan.DataSource = rs
            dgTipePekerjaan.Columns(0).Width = 750
            dgTipePekerjaan.Columns(0).Alignment = vbCenter
            dgTipePekerjaan.Columns(1).Width = 4200
            dgTipePekerjaan.ReBind
            Set rs = Nothing

        Case 3 'eselon
            Set rs = Nothing
            strSQL = "select * from Eselon"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgEselon.DataSource = rs
            dgEselon.Columns(0).Width = 750
            dgEselon.Columns(0).Alignment = vbCenter
            dgEselon.Columns(1).Width = 4200
            dgEselon.ReBind
            Set rs = Nothing

        Case 4 'komponen gaji
            Set rs = Nothing
            strSQL = "select * from komponengaji"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgKomponenGaji.DataSource = rs
            dgKomponenGaji.Columns(0).Width = 750
            dgKomponenGaji.Columns(0).Alignment = vbCenter
            dgKomponenGaji.Columns(1).Width = 4200
            dgKomponenGaji.ReBind
            Set rs = Nothing

        Case 5 'jenis hukuman
            Set rs = Nothing
            strSQL = "select * from JenisHukuman"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgJenisHukum.DataSource = rs
            dgJenisHukum.Columns(0).Width = 750
            dgJenisHukum.Columns(0).Alignment = vbCenter
            dgJenisHukum.Columns(1).Width = 4200
            dgJenisHukum.ReBind
            Set rs = Nothing

        Case 6 'Status perkawian
            Set rs = Nothing
            strSQL = "select * from StatusPerkawinan"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgStatusPerkawinan.DataSource = rs
            dgStatusPerkawinan.Columns(0).Width = 750
            dgStatusPerkawinan.Columns(0).Alignment = vbCenter
            dgStatusPerkawinan.Columns(1).Width = 4200
            dgStatusPerkawinan.ReBind
            Set rs = Nothing

        Case 7 'Agama
            Set rs = Nothing
            strSQL = "select * from Agama"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgAgama.DataSource = rs
            dgAgama.Columns(0).Width = 750
            dgAgama.Columns(0).Alignment = vbCenter
            dgAgama.Columns(1).Width = 4200
            dgAgama.ReBind
            Set rs = Nothing

        Case 8 'Golongan darah
            Set rs = Nothing
            strSQL = "select * from GolonganDarah"
            rs.Open strSQL, dbConn, adOpenForwardOnly, adLockReadOnly
            Set dgGolonganDarah.DataSource = rs
            dgGolonganDarah.Columns(0).Width = 750
            dgGolonganDarah.Columns(0).Alignment = vbCenter
            dgGolonganDarah.Columns(1).Width = 4200
            dgGolonganDarah.ReBind
            Set rs = Nothing
    End Select
    Call cmdBatal_Click
End Sub

Private Sub txtNmExt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtAgama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtEselon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtGaji_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtHukuman_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtKawin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtNmExtTipe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub

Private Sub txtStatusPerkawinan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtKawin.SetFocus
End Sub

Private Sub txtTipePekerjaan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtKdExtTipe.SetFocus
End Sub

Private Sub txtKdExtJnsDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkStsJnsDiklat.SetFocus
End Sub

Private Sub chkStsJnsDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNmExtJnsDiklat.SetFocus
End Sub

Private Sub txtNmExtJnsDiklat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSimpan.SetFocus
End Sub
