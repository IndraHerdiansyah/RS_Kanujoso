VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmKPIN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Minta PIN"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3675
   Icon            =   "frmKPIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   3675
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      GridColorFixed  =   12632256
      GridLinesFixed  =   1
   End
   Begin VB.CommandButton cmdAmbil 
      Caption         =   "&Ambil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdBatal 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
      Begin VB.TextBox txtAdd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Alamat FRS-400 :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   6
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
      Height          =   945
      Left            =   2400
      Picture         =   "frmKPIN.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKPIN.frx":1A52
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4935
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   0
      Picture         =   "frmKPIN.frx":30B0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
End
Attribute VB_Name = "frmKPIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAmbil_Click()

Dim pinAmbil As String

With frmPINAbsensiPegawai
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = "" Then
    GoTo hell
End If
    
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = True Then

i = MsgBox("Apakah No. PIN " & MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) & " Akan Diambil ?", vbOKCancel, "Pesan Ambil PIN")

    If i = vbOK Then
    
        '02 0D 01 11 30 30 30 30 30 34 31 33 03 1A
        If frmAbsensiPegawai.MSComm1.PortOpen = True Then
            With frmAbsensiPegawai
                fp = 0
                .timerPIN.Enabled = False
                .TimerHapusPIN.Enabled = False
                .minta_absensi.Enabled = False
                pinAmbil = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
                pinAmbil = "00000000" & pinAmbil
                pinAmbil = Right(pinAmbil, 8)
                h11 = Asc(Mid(pinAmbil, 1, 1))
                h12 = Asc(Mid(pinAmbil, 2, 1))
                h13 = Asc(Mid(pinAmbil, 3, 1))
                h14 = Asc(Mid(pinAmbil, 4, 1))
                h15 = Asc(Mid(pinAmbil, 5, 1))
                h16 = Asc(Mid(pinAmbil, 6, 1))
                h17 = Asc(Mid(pinAmbil, 7, 1))
                h18 = Asc(Mid(pinAmbil, 8, 1))
                .MSComm1.Output = Chr$(&H2) & Chr$(&HD) & Chr$(&H1) & Chr$(&H11) & Chr$(h11) & Chr$(h12) & Chr$(h13) & Chr$(h14) & Chr$(h15) & Chr$(h16) & Chr$(h17) & Chr$(h18) & Chr$(&H3) & Chr$(&H1A)
            End With
            .txtNoPIN.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col)
            MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.Col) = Remove
            .txttanggal.Text = Now
        End If
    End If
    
End If

End With

hell:

End Sub

Private Sub cmdBatal_Click()
With frmAbsensiPegawai
    If .MSComm1.PortOpen = True Then
        .timerPIN.Enabled = False
        .minta_absensi.Enabled = True
        a = 0
        hasil = 0
        jumlah = 0
        simpanPIN = ""
    End If
End With
Unload Me
End Sub

Private Sub cmdConnect_Click()
cmdConnect.Enabled = False
cmdAmbil.Enabled = False
If frmAbsensiPegawai.MSComm1.PortOpen = True Then
    If txtAdd.Text = "" Then
        cmdConnect.Enabled = True
        cmdAmbil.Enabled = True
        MsgBox "Alamat FRS-400 Kosong", vbOKOnly, "Pesan koneksi"
    ElseIf txtAdd.Text > add2 Or txtAdd.Text <= 0 Then
        cmdConnect.Enabled = True
        cmdAmbil.Enabled = True
        MsgBox "Alamat FRS-400 Tujuan Salah", vbOKOnly, "Pesan koneksi"
    Else
        addpin = txtAdd.Text
        With frmAbsensiPegawai
            If .minta_absensi.Enabled = True Then
                .minta_absensi.Enabled = False
                .TimerHapusPIN.Enabled = False
                .timerPIN.Enabled = True
            End If
        End With
    End If
ElseIf frmAbsensiPegawai.MSComm1.PortOpen = False Then
    cmdConnect.Enabled = True
    cmdAmbil.Enabled = True
    MsgBox "Tidak Ada Koneksi dengan FRS-400", vbOKOnly, "Pesan koneksi"
End If
End Sub

Private Sub Form_Load()
    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)
    MSFlexGrid1.ColWidth(1) = 2200
End Sub
