VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRekapLaporanGajiNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirt2000 - Rekap Laporan Gaji Pegawai"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   Icon            =   "frmRekapLaporanGaji.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   7335
         Begin VB.CommandButton cmdTutup 
            Caption         =   "Tutup"
            Height          =   375
            Left            =   5160
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
         Begin VB.CommandButton cmdCetak 
            Caption         =   "Cetak"
            Height          =   375
            Left            =   3120
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   7335
         Begin VB.CheckBox chkPerbagian 
            Caption         =   "Perbagian / Ruangan"
            Height          =   195
            Left            =   5160
            TabIndex        =   15
            Top             =   480
            Width           =   1935
         End
         Begin VB.Frame Frame4 
            Caption         =   "Periode"
            Height          =   735
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   4935
            Begin MSComCtl2.DTPicker dtpTglAwal 
               Height          =   330
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
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
               CustomFormat    =   "yyyy"
               Format          =   116326400
               UpDown          =   -1  'True
               CurrentDate     =   38448
            End
            Begin MSComCtl2.DTPicker dtpTglAhir 
               Height          =   330
               Left            =   2520
               TabIndex        =   6
               Top             =   240
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   582
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
               CustomFormat    =   "yyyy"
               Format          =   116326400
               UpDown          =   -1  'True
               CurrentDate     =   38448
            End
         End
         Begin MSDataListLib.DataCombo dcJabatan 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcPangkat 
            Height          =   315
            Left            =   2520
            TabIndex        =   8
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcRuangan 
            Height          =   315
            Left            =   4920
            TabIndex        =   14
            Top             =   1200
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label Label1 
            Caption         =   "Bagian / Ruangan"
            Height          =   255
            Left            =   4920
            TabIndex        =   12
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Jabatan"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Pangkat"
            Height          =   255
            Left            =   2520
            TabIndex        =   9
            Top             =   960
            Width           =   1335
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
      Left            =   5880
      Picture         =   "frmRekapLaporanGaji.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmRekapLaporanGaji.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmRekapLaporanGajiNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dgRekapLaporan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    'Call subLoadGridTab
End Sub

Private Sub chkPerbagian_Click()
    If chkPerbagian.Value = 1 Then
        dcRuangan.Enabled = True
        Call msubDcSource(dcRuangan, rs, "SELECT DISTINCT kdRuangan, NamaRuangan FROM V_RuanganInstalasi where NamaRuangan LIKE '%" & dcRuangan.Text & "%' and StatusEnabled = 1 ORDER BY NamaRuangan")
        If rs.EOF = False Then dcRuangan.BoundText = rs(0).Value
    Else
        dcRuangan.Enabled = False
        dcRuangan.Text = ""
    End If
End Sub

Private Sub cmdCetak_Click()
    frmCetakRekapLaporanGaji.Show
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    dtpTglAwal.Value = Format(Now, "yyyy/MMMM/dd")
    dtpTglAhir.Value = Format(Now, "yyyy/MMMM/dd")
    Call PlayFlashMovie(Me)
    'Call subLoadGridTab
    Call subLoadDcSource
End Sub

Public Sub subLoadDcSource()
    On Error GoTo errLoad
    Call msubDcSource(dcJabatan, rs, "SELECT KdJabatan, NamaJabatan FROM Jabatan where NamaJabatan LIKE '%" & dcJabatan.Text & "%' and StatusEnabled = 1 ORDER BY NamaJabatan")
    If rs.EOF = False Then dcJabatan.BoundText = rs(0).Value
    Call msubDcSource(dcPangkat, rs, "SELECT KdPangkat, NamaPangkat FROM Pangkat where NamaPangkat LIKE '%" & dcPangkat.Text & "%' and StatusEnabled = 1 ORDER BY NamaPangkat")
    If rs.EOF = False Then dcPangkat.BoundText = rs(0).Value
    Exit Sub
errLoad:
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'Call subLoadGridTab
End Sub
