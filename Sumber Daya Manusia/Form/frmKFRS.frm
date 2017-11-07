VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash8.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmKFRS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medifirst 2000 - Hasil Koneksi"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   Icon            =   "frmKFRS.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6495
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1931
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "&BATAL"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      TabIndex        =   0
      Top             =   2640
      Width           =   1335
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash Flash1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
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
      Height          =   975
      Left            =   0
      Picture         =   "frmKFRS.frx":0CCA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   945
      Left            =   4920
      Picture         =   "frmKFRS.frx":368B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   1800
      Picture         =   "frmKFRS.frx":4413
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   3120
      Picture         =   "frmKFRS.frx":5A71
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "FRS-400 yang terkoneksi :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "frmKFRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    With frmAbsensiPegawai
        .List1.List(0) = "Status FRS-400 : "
        For i = 1 To add2
            .List1.List(i) = "       " & "FRS - " & i & ""
        Next i
        fRS = &H0
        inbuff = ""
        inbuff2 = ""
        .minta_absensi.Enabled = True
        .tmr_CekError.Enabled = True

    End With
    Unload Me
End Sub

Private Sub cmdBatal_Click()
    frmAbsensiPegawai.cmdsetting.Enabled = True
    If frmAbsensiPegawai.MSComm1.PortOpen = True Then
        frmAbsensiPegawai.MSComm1.PortOpen = False
    End If
    add2 = &H0
    Unload Me

End Sub

Private Sub Form_Load()

    Call centerForm(Me, MDIUtama)
    Call PlayFlashMovie(Me)

    With MSFlexGrid1
        .ColWidth(0) = 1000
        .ColWidth(1) = 2450
        .ColWidth(2) = 2450
        .Width = 6275
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    With frmAbsensiPegawai

        If .minta_absensi.Enabled = False Then
            .cmdConnect.Enabled = True
            .cmdDisconnect.Enabled = False
        End If

    End With

    a = 0

End Sub
