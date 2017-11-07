VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMonitoringJadwalKerjaRuangan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Monitoring Jadwal Kerja Ruangan"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12405
   Icon            =   "frmMonitoringJadwalKerjaRuangan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   12405
   Begin VB.Frame Frame3 
      Caption         =   "Parameter"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   12135
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
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
         Left            =   11040
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin MSDataListLib.DataCombo dcRuangan 
         Height          =   330
         Left            =   3120
         TabIndex        =   9
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.TextBox txtNama 
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
         Height          =   330
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpTglAwal 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   107806723
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin MSComCtl2.DTPicker dtpTglAkhir 
         Height          =   375
         Left            =   8640
         TabIndex        =   13
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
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
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   107806723
         UpDown          =   -1  'True
         CurrentDate     =   38209
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "s/d"
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
         Left            =   8280
         TabIndex        =   14
         Top             =   555
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Ruangan"
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
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Nama Pegawai"
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
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   7200
      Width           =   12135
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
         Height          =   495
         Left            =   7560
         TabIndex        =   11
         Top             =   240
         Width           =   2175
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
         Height          =   495
         Left            =   9840
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   12135
      Begin MSDataGridLib.DataGrid dgMonitoringJadwalKerja 
         Height          =   4335
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   7646
         _Version        =   393216
         HeadLines       =   2
         RowHeight       =   16
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
      Height          =   975
      Left            =   10200
      Picture         =   "frmMonitoringJadwalKerjaRuangan.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmMonitoringJadwalKerjaRuangan.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmMonitoringJadwalKerjaRuangan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBatal_Click()
    txtNama.Text = ""
    dcRuangan.Text = ""
    Call subLoadDataPegawai
End Sub

Private Sub cmdCari_Click()
    Call subLoadDataPegawai
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dcRuangan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then cmdCari.SetFocus
End Sub

Private Sub dtpTglAkhir_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdCari.SetFocus
    End If
End Sub

Private Sub dtpTglAwal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        dtpTglAkhir.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    dtpTglAkhir.Value = Now
    dtpTglAwal.Value = Now
    
    Call subDCSource
    Call subLoadDataPegawai
End Sub

Private Sub subLoadDataPegawai()
    On Error GoTo errLoad
    strSQL = "Select * from V_JadwalKerjaNew where Ruangan like '%" & dcRuangan.Text & "%' and Nama like '%" & txtNama.Text & "%' and(Tanggal between '" & Format(dtpTglAwal.Value, "yyyy/MM/dd 00:00:00") & "' and '" & Format(dtpTglAkhir.Value, "yyyy/MM/dd 23:59:59") & "')"
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgMonitoringJadwalKerja.DataSource = rsb
    With dgMonitoringJadwalKerja
        .Columns(0).Width = 1300
        .Columns(1).Width = 3500
        .Columns(2).Width = 2500
        .Columns(3).Width = 0
        .Columns(4).Width = 0
        .Columns(5).Width = 2000
        .Columns(6).Width = 2000
'        .Columns(7).Width = 0'//yna 2014-0808
'        .Columns(0).Caption = "ID Pegawai"
'        .Columns(2).Caption = "Jenis Pegawai"
'        .Columns(4).Caption = "Kelompok Pegawai"
'        .Columns(5).Caption = "Nama Lengkap"
'        .Columns(6).Caption = "Jenis Kelamin"
    End With
    Exit Sub
errLoad:
    Call msubPesanError
End Sub

Sub subDCSource()
    strSQL = "SELECT KdRuangan, NamaRuangan FROM Ruangan order by NamaRuangan"
    Call msubDcSource(dcRuangan, rs, strSQL)
End Sub

Private Sub txtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    If KeyAscii = 13 Then dcRuangan.SetFocus
End Sub
