VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash10c.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmNotifikasi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst2000 - Notifikasi"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   14475
   Begin VB.CommandButton Command3 
      Caption         =   "Cetak"
      Height          =   375
      Left            =   11760
      TabIndex        =   12
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   9840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   4455
   End
   Begin VB.ComboBox Combo2 
      Height          =   330
      Left            =   960
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   120
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   5880
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ingatkan"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cari"
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Tutu&p"
      Height          =   375
      Left            =   13080
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   5741
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
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpAwal 
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Format          =   16515075
      CurrentDate     =   37760
   End
   Begin MSComCtl2.DTPicker dtpAkhir 
      Height          =   330
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
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
      Format          =   16515075
      CurrentDate     =   37760
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   2760
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _cx             =   3175
      _cy             =   1720
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   "-1"
      Loop            =   "-1"
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   "0"
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   "0"
      EmbedMovie      =   "0"
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   "1"
      Profile         =   "0"
      ProfileAddress  =   ""
      ProfilePort     =   "0"
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Image Image3 
      Height          =   975
      Left            =   0
      Picture         =   "frmNotifikasi.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmNotifikasi.frx":29C1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Desc"
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Periode :"
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "s/d"
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
End
Attribute VB_Name = "frmNotifikasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    strSQL = "select IdPegawai, NamaLengkap, NamaJenjangJabatan, Deskripsi, Desk,tgl as 'Tanggal Periode',NamaRuangan, Tanggal from V_Notifikasi where tgl between '" & Format(dtpAwal.Value, "yyyy-MM-dd") & "' and '" & Format(dtpAkhir.Value, "yyyy-MM-dd") & "' " & _
             " and Deskripsi like '%" & Combo2.Text & "%'"
    Call msubRecFO(rs, strSQL)
    Set dg.DataSource = rs
    
    dg.Columns(1).Width = 3000
End Sub

Private Sub Command2_Click()
    Select Case Combo1.Text
        Case "1 Jam lagi"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~1"
        Case "5 Jam lagi"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~5"
        Case "10 Jam lagi"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~10"
        Case "Lupakan"
            SaveSetting "SDM", "Notif", "Ultah", Date & "~0"
    End Select
    
    MsgBox "OKE", vbInformation, "..:."
End Sub

Private Sub Command3_Click()
On Error GoTo errLoad
Dim pesan As VbMsgBoxResult
    
     strSQL = "select IdPegawai, NamaLengkap, NamaJenjangJabatan, Deskripsi, Desk,tgl as 'Tanggal Periode',NamaRuangan, Tanggal from V_Notifikasi where tgl between '" & Format(dtpAwal.Value, "yyyy-MM-dd") & "' and '" & Format(dtpAkhir.Value, "yyyy-MM-dd") & "' " & _
             " and Deskripsi like '%" & Combo2.Text & "%'"

    pesan = MsgBox("Apakah anda ingin langsung mencetak laporan? " & vbNewLine & "Pilih No jika ingin ditampilkan terlebih dahulu ", vbQuestion + vbYesNo, "Konfirmasi")
    vLaporan = ""
    If pesan = vbYes Then vLaporan = "Print"

    frm_cetak_LaporanNotifikasi.Show
    Exit Sub
errLoad:


End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    
    dtpAwal.Value = Now()
    dtpAkhir.Value = Now()
    
    Combo1.AddItem "1 Jam lagi"
    Combo1.AddItem "5 Jam lagi"
    Combo1.AddItem "10 Jam lagi"
    Combo1.AddItem "Lupakan"
    Combo1.Text = ""
    
    Combo2.AddItem "Kenaikan Pangkat"
    Combo2.AddItem "Ulang tahun"
    Combo2.AddItem "Pensiun"
    Combo2.AddItem "Habis Kontrak"
    Combo2.Text = ""
    
    strSQL = "select * from V_Notifikasi where tgl ='" & Format(dtpAwal.Value, "yyyy-MM-dd") & "'"
    Call msubRecFO(rs, strSQL)
    Set dg.DataSource = rs
End Sub
