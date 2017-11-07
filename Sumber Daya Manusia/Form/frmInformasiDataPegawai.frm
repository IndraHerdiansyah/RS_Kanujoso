VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmInformasiDataPegawai 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Informasi Data Pegawai"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   Icon            =   "frmInformasiDataPegawai.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11070
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
      Height          =   330
      Left            =   9240
      TabIndex        =   11
      Top             =   7200
      Width           =   1575
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
      Height          =   330
      Left            =   7560
      TabIndex        =   10
      Top             =   7200
      Width           =   1575
   End
   Begin VB.OptionButton optJenis 
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
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.OptionButton optNama 
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
      Height          =   255
      Left            =   6840
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.OptionButton optID 
      Caption         =   "ID Pegawai"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameter Pencarian"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   10815
      Begin VB.Frame Frame2 
         Caption         =   "Kriteria Pencarian"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   6
         Top             =   360
         Width           =   6375
      End
      Begin VB.TextBox txtParameter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   0
         Top             =   630
         Width           =   2895
      End
      Begin VB.CommandButton cmdCari 
         Caption         =   "&Cari"
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
         Left            =   3240
         TabIndex        =   1
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Masukkan Parameter"
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1680
      End
   End
   Begin MSDataGridLib.DataGrid dgPegawai 
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8493
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   16
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   9
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
      Picture         =   "frmInformasiDataPegawai.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   9240
      Picture         =   "frmInformasiDataPegawai.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmInformasiDataPegawai.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13095
   End
End
Attribute VB_Name = "frmInformasiDataPegawai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilter As String
Dim rsb As New ADODB.recordset

Private Sub subLoadDataPasien()
    On Error GoTo Errload
    strSQL = "Select * from V_DataPegawai " & strFilter
    Set rsb = Nothing
    rsb.Open strSQL, dbConn, adOpenStatic, adLockOptimistic
    Set dgPegawai.DataSource = rsb
    With dgPegawai
        .Columns(0).Width = 1300
        .Columns(1).Width = 0
        .Columns(2).Width = 2050
        .Columns(3).Width = 0
        .Columns(4).Width = 2500
        .Columns(5).Width = 3200
        .Columns(6).Width = 1500
        .Columns(7).Width = 2000
        .Columns(8).Width = 1500
        .Columns(9).Width = 2500
        .Columns(0).Caption = "ID Pegawai"
        .Columns(2).Caption = "Jenis Pegawai"
        .Columns(4).Caption = "Kelompok Pegawai"
        .Columns(5).Caption = "Nama Lengkap"
        .Columns(6).Caption = "Jenis Kelamin"
        .Columns(7).Caption = "Pangkat"
        .Columns(8).Caption = "Golongan"
        .Columns(9).Caption = "Jabatan"
    End With
    Exit Sub
Errload:
    Call msubPesanError
End Sub

Private Sub cmdBaru_Click()
    strFilter = "ORDER by IdPegawai"
    Call subLoadDataPasien
    txtParameter.Text = ""
    txtParameter.SetFocus
End Sub

Public Sub cmdCari_Click()
    If optID.Value = True Then
        strFilter = " WHERE IdPegawai like '%" & txtParameter.Text & "%'"
    ElseIf optNama.Value = True Then
        strFilter = " WHERE NamaLengkap like '%" & txtParameter.Text & "%'"
    ElseIf optJenis.Value = True Then
        strFilter = " WHERE JenisPegawai like '%" & txtParameter.Text & "%'"
    End If

    Call subLoadDataPasien
    If rsb.RecordCount = 0 Then Exit Sub
    dgPegawai.SetFocus
End Sub

Private Sub cmdTutup_Click()
    Unload Me
End Sub

Private Sub dgPegawai_Click()
    WheelHook.WheelUnHook
    Set MyProperty = dgPegawai
    WheelHook.WheelHook dgPegawai
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    optID.Value = True
    strFilter = "ORDER by NamaLengkap ASC"
    Call subLoadDataPasien

End Sub

Private Sub Form_Unload(Cancel As Integer)
    blnFrmCariPasien = False
End Sub

Private Sub txtParameter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    
    If optNama.Value = True Or optJenis.Value = True Then
        Call SetKeyPressToChar(KeyAscii)
    End If
    
    If KeyAscii = 13 Then
        Call cmdCari_Click
    End If
End Sub
